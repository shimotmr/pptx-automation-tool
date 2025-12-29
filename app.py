import streamlit as st
import streamlit.components.v1 as components
import os
import uuid
import json
import shutil
import traceback
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# ==============================
# 基本設定
# ==============================
st.set_page_config(
    page_title="Aurotek｜簡報案例自動化發布平台",
    page_icon="📊",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# ==============================
# 樣式
# ==============================
st.markdown("""
<style>
header[data-testid="stHeader"] { display: none; }
.block-container { padding-top: 1rem; }

.callout{
  border:1px solid #E5E7EB;
  border-radius:14px;
  padding:14px;
  margin:10px 0;
  background:#F8FAFC;
}
.callout.blue{
  border-left:6px solid #0B4F8A;
  background:#EAF3FF;
  color:#0B4F8A;
  font-weight:700;
}

.section{
  border:1px solid #E5E7EB;
  border-radius:16px;
  padding:16px;
  background:#fff;
}
</style>
""", unsafe_allow_html=True)

# ==============================
# Helper
# ==============================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)

def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)
    os.makedirs(WORK_DIR, exist_ok=True)

def load_history(filename):
    if not os.path.exists(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get(filename, [])
    except:
        return []

def save_history(filename, jobs):
    data = {}
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except:
            data = {}
    data[filename] = jobs
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ==============================
# Header
# ==============================
st.markdown(f"""
<div style="text-align:center;margin-bottom:10px;">
  <img src="{LOGO_URL}" style="width:300px;" />
  <div style="letter-spacing:2px;font-weight:600;color:#6B7280;">
    簡報案例自動化發布平台
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="callout blue">
上傳簡報 → 拆分任務 → 影片雲端化 → 內嵌優化 → Google Slides 發布 → 寫入資料庫
</div>
""", unsafe_allow_html=True)

# ==============================
# 初始化狀態
# ==============================
if "split_jobs" not in st.session_state:
    st.session_state.split_jobs = []

if "current_file" not in st.session_state:
    st.session_state.current_file = None

if "ppt_meta" not in st.session_state:
    st.session_state.ppt_meta = {"total": 0}

if "bot" not in st.session_state:
    st.session_state.bot = PPTAutomationBot()

# ==============================
# Step 1：上傳檔案
# ==============================
ensure_workspace()
source_path = os.path.join(WORK_DIR, "source.pptx")

st.markdown('<div class="section">', unsafe_allow_html=True)
st.subheader("步驟一：上傳簡報")

uploaded = st.file_uploader("PPTX", type=["pptx"])
if uploaded:
    if st.session_state.current_file != uploaded.name:
        cleanup_workspace()
        with open(source_path, "wb") as f:
            f.write(uploaded.getbuffer())

        prs = Presentation(source_path)
        st.session_state.ppt_meta["total"] = len(prs.slides)
        st.session_state.split_jobs = load_history(uploaded.name)
        st.session_state.current_file = uploaded.name

    st.markdown(
        f"<div class='callout blue'>已讀取 {uploaded.name}（共 {st.session_state.ppt_meta['total']} 頁）</div>",
        unsafe_allow_html=True
    )
st.markdown("</div>", unsafe_allow_html=True)

# ==============================
# Step 2：拆分任務（完整欄位）
# ==============================
if st.session_state.current_file:
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.subheader("步驟二：設定拆分任務")

    if st.button("新增任務"):
        st.session_state.split_jobs.append({
            "id": str(uuid.uuid4()),
            "filename": "",
            "start": 1,
            "end": 1,
            "category": "",
            "sub_category": "",
            "client": "",
            "keywords": ""
        })

    for job in st.session_state.split_jobs:
        with st.container(border=True):
            c1, c2, c3 = st.columns([3,1,1])
            job["filename"] = c1.text_input("檔名", job["filename"], key=f"f_{job['id']}")
            job["start"] = c2.number_input("起始頁", 1, st.session_state.ppt_meta["total"], job["start"], key=f"s_{job['id']}")
            job["end"] = c3.number_input("結束頁", 1, st.session_state.ppt_meta["total"], job["end"], key=f"e_{job['id']}")

            c4, c5, c6, c7 = st.columns(4)
            job["category"] = c4.text_input("類型", job["category"], key=f"cat_{job['id']}")
            job["sub_category"] = c5.text_input("子分類", job["sub_category"], key=f"sub_{job['id']}")
            job["client"] = c6.text_input("客戶", job["client"], key=f"cli_{job['id']}")
            job["keywords"] = c7.text_input("關鍵字", job["keywords"], key=f"kw_{job['id']}")

    save_history(st.session_state.current_file, st.session_state.split_jobs)
    st.markdown("</div>", unsafe_allow_html=True)

# ==============================
# Step 3：執行
# ==============================
if st.session_state.current_file:
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.subheader("步驟三：開始執行")

    progress = st.progress(0)
    status = st.empty()

    if st.button("執行自動化排程"):
        try:
            def update(step, pct):
                progress.progress(pct)
                status.markdown(
                    f"<div class='callout blue'>步驟 {step} 進行中（{pct}%）</div>",
                    unsafe_allow_html=True
                )

            update("1/5 影片處理", 10)
            video_map = st.session_state.bot.extract_and_upload_videos(source_path)

            update("2/5 影片置換", 30)
            mod_path = os.path.join(WORK_DIR, "mod.pptx")
            st.session_state.bot.replace_videos_with_images(
                source_path, mod_path, video_map
            )

            update("3/5 檔案優化", 50)
            slim_path = os.path.join(WORK_DIR, "slim.pptx")
            st.session_state.bot.shrink_pptx(mod_path, slim_path)

            update("4/5 拆分上傳", 70)
            results = st.session_state.bot.split_and_upload(
                slim_path, st.session_state.split_jobs
            )

            update("5/5 寫入資料庫", 90)
            st.session_state.bot.log_to_sheets(results)

            update("完成", 100)

            st.markdown("<div class='callout blue'>流程完成</div>", unsafe_allow_html=True)

            # ===== 完成圖卡 =====
            st.subheader("產出結果")
            for r in results:
                with st.container(border=True):
                    st.markdown(f"**{r['filename']}**")
                    c1, c2 = st.columns(2)
                    c1.link_button("開啟簡報", r["final_link"])
                    c2.code(r["final_link"])

        except Exception as e:
            st.error(f"流程失敗：{e}")
            st.code(traceback.format_exc())

    st.markdown("</div>", unsafe_allow_html=True)
