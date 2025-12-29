import streamlit as st
import streamlit.components.v1 as components
import os
import uuid
import json
import shutil
import traceback
import requests
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# =========================================================
# 基本設定
# =========================================================
st.set_page_config(
    page_title="Aurotek｜簡報案例自動化發布平台",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# =========================================================
# CSS（保留你目前企業版風格）
# =========================================================
st.markdown("""
<style>
header[data-testid="stHeader"]{display:none;}
.block-container{padding-top:0.8rem;}

:root{
  --brand:#0B4F8A;
  --brand-bg:#EAF3FF;
  --border:#E5E7EB;
  --text:#111827;
  --muted:#6B7280;
}

.auro-header{
  display:flex;
  flex-direction:column;
  align-items:center;
  margin-bottom:8px;
}
.auro-header img{width:300px;height:auto;}
.auro-sub{color:var(--muted);font-weight:600;letter-spacing:2px;}

.callout{
  border:1px solid var(--border);
  border-left:4px solid var(--brand);
  background:var(--brand-bg);
  padding:12px 14px;
  border-radius:12px;
  margin:10px 0;
  font-weight:650;
}

.section{
  border:1px solid var(--border);
  border-radius:16px;
  padding:14px;
  margin-bottom:16px;
  background:#fff;
}

.stProgress > div > div > div > div{color:#fff;font-weight:600;}
</style>
""", unsafe_allow_html=True)

# =========================================================
# Helper
# =========================================================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)

def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR, ignore_errors=True)
    os.makedirs(WORK_DIR, exist_ok=True)

def load_history(filename):
    if not os.path.exists(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f).get(filename, [])
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

def add_job(total):
    st.session_state.jobs.append({
        "id": str(uuid.uuid4())[:8],
        "filename": "",
        "start": 1,
        "end": total,
        "category": "清潔",
        "subcategory": "",
        "client": "",
        "keywords": ""
    })

def validate_jobs(jobs, total):
    errs = []
    for j in jobs:
        if not j["filename"]:
            errs.append("檔名不可空白")
        if j["start"] > j["end"]:
            errs.append("起始頁不可大於結束頁")
        if j["end"] > total:
            errs.append("頁數超出總頁數")
    return errs

# =========================================================
# 🔒 關鍵：安全取代影片（無影片也不中斷）
# =========================================================
def safe_replace_videos(bot, source, out_path, video_map):
    """
    無影片時：
    - 不呼叫 replace_videos_with_images
    - 直接複製 source → out_path
    """
    if not video_map:
        shutil.copyfile(source, out_path)
        return

    bot.replace_videos_with_images(
        source,
        out_path,
        video_map,
        progress_callback=lambda c, t: None
    )

# =========================================================
# Header
# =========================================================
st.markdown(f"""
<div class="auro-header">
  <img src="{LOGO_URL}">
  <div class="auro-sub">簡報案例自動化發布平台</div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="callout">
流程：上傳簡報 → 拆分任務 → 影片處理（可略） → 拆分發布 → 寫入資料庫
</div>
""", unsafe_allow_html=True)

# =========================================================
# Init session
# =========================================================
if "jobs" not in st.session_state:
    st.session_state.jobs = []
if "current_file" not in st.session_state:
    st.session_state.current_file = None
if "total_slides" not in st.session_state:
    st.session_state.total_slides = 0
if "bot" not in st.session_state:
    st.session_state.bot = PPTAutomationBot()

ensure_workspace()
SOURCE_PATH = os.path.join(WORK_DIR, "source.pptx")

# =========================================================
# Step 1
# =========================================================
with st.container():
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.subheader("步驟一：選擇簡報")

    uploaded = st.file_uploader("PPTX", type=["pptx"], label_visibility="collapsed")
    if uploaded:
        if st.session_state.current_file != uploaded.name:
            cleanup_workspace()
            with open(SOURCE_PATH, "wb") as f:
                f.write(uploaded.getbuffer())

            prs = Presentation(SOURCE_PATH)
            st.session_state.total_slides = len(prs.slides)
            st.session_state.current_file = uploaded.name
            st.session_state.jobs = load_history(uploaded.name)

        st.markdown(
            f"<div class='callout'>已讀取 {uploaded.name}（共 {st.session_state.total_slides} 頁）</div>",
            unsafe_allow_html=True
        )

    st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# Step 2
# =========================================================
if st.session_state.current_file:
    with st.container():
        st.markdown("<div class='section'>", unsafe_allow_html=True)
        st.subheader("步驟二：設定拆分任務")

        if st.button("新增任務"):
            add_job(st.session_state.total_slides)

        for i, job in enumerate(st.session_state.jobs):
            with st.expander(f"任務 {i+1}", expanded=True):
                c1, c2, c3 = st.columns([3,1,1])
                job["filename"] = c1.text_input("檔名", job["filename"], key=f"f{i}")
                job["start"] = c2.number_input("起始頁", 1, st.session_state.total_slides, job["start"], key=f"s{i}")
                job["end"] = c3.number_input("結束頁", 1, st.session_state.total_slides, job["end"], key=f"e{i}")

                m1, m2, m3, m4 = st.columns(4)
                job["category"] = m1.text_input("類型", job["category"])
                job["subcategory"] = m2.text_input("子分類", job["subcategory"])
                job["client"] = m3.text_input("客戶", job["client"])
                job["keywords"] = m4.text_input("關鍵字", job["keywords"])

        save_history(st.session_state.current_file, st.session_state.jobs)
        st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# Step 3
# =========================================================
if st.session_state.current_file:
    with st.container():
        st.markdown("<div class='section'>", unsafe_allow_html=True)
        st.subheader("步驟三：開始執行")

        if st.button("執行自動化排程", use_container_width=True):
            errs = validate_jobs(st.session_state.jobs, st.session_state.total_slides)
            if errs:
                for e in errs:
                    st.error(e)
                st.stop()

            bot = st.session_state.bot
            main = st.progress(0)

            try:
                # Step 1：影片
                main.progress(10, "檢查影片")
                video_map = bot.extract_and_upload_videos(
                    SOURCE_PATH,
                    os.path.join(WORK_DIR, "media"),
                    file_prefix=os.path.splitext(st.session_state.current_file)[0],
                    progress_callback=lambda f,c,t: None,
                    log_callback=lambda x: None
                ) or {}

                if not video_map:
                    st.markdown("<div class='callout'>未偵測到影片，略過影片處理</div>", unsafe_allow_html=True)

                # Step 2：replace
                main.progress(30, "處理簡報")
                MOD_PATH = os.path.join(WORK_DIR, "modified.pptx")
                safe_replace_videos(bot, SOURCE_PATH, MOD_PATH, video_map)

                # Step 3：shrink
                main.progress(45, "壓縮優化")
                SLIM_PATH = os.path.join(WORK_DIR, "slim.pptx")
                bot.shrink_pptx(MOD_PATH, SLIM_PATH, progress_callback=lambda c,t: None)

                # Step 4：split
                main.progress(65, "拆分並上傳")
                results = bot.split_and_upload(
                    SLIM_PATH,
                    st.session_state.jobs,
                    file_prefix=os.path.splitext(st.session_state.current_file)[0],
                    progress_callback=lambda f,c,t: None,
                    log_callback=lambda x: None
                )

                if not results:
                    raise RuntimeError("拆分後沒有產出任何結果")

                # Step 5：embed
                main.progress(85, "嵌入影片")
                final = bot.embed_videos_in_slides(results, log_callback=lambda x: None)

                # Step 6：log
                main.progress(95, "寫入資料庫")
                bot.log_to_sheets(final, log_callback=lambda x: None)

                main.progress(100, "完成")
                st.markdown("<div class='callout'>流程完成</div>", unsafe_allow_html=True)

            except Exception as e:
                st.error(str(e))
                st.code(traceback.format_exc())

        st.markdown("</div>", unsafe_allow_html=True)
