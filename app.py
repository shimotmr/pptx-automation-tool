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
    page_icon="📊",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# =========================================================
# 工具函式
# =========================================================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)

def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR, ignore_errors=True)
    os.makedirs(WORK_DIR, exist_ok=True)

def detect_resume_step():
    """
    斷點續傳判斷：
    1 = 從頭
    2 = 已有 source.pptx
    3 = 已有 modified.pptx
    4 = 已有 slim.pptx
    """
    if os.path.exists(os.path.join(WORK_DIR, "slim.pptx")):
        return 4
    if os.path.exists(os.path.join(WORK_DIR, "modified.pptx")):
        return 3
    if os.path.exists(os.path.join(WORK_DIR, "source.pptx")):
        return 2
    return 1

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

def reset_to_step1():
    for k in [
        "current_file_name",
        "ppt_meta",
        "split_jobs",
    ]:
        if k in st.session_state:
            del st.session_state[k]
    cleanup_workspace()
    st.rerun()

# =========================================================
# Header（HTML，LOGO 鎖 300px）
# =========================================================
components.html(f"""
<div style="display:flex;flex-direction:column;align-items:center;margin-bottom:6px;">
  <img id="auro-logo" src="{LOGO_URL}" style="width:300px;max-width:90vw;height:auto;" />
  <div style="margin-top:4px;font-size:1rem;font-weight:600;letter-spacing:2px;color:#6B7280;">
    簡報案例自動化發布平台
  </div>
</div>

<style>
@media (max-width:768px){{
  #auro-logo {{ width:260px !important; }}
}}
</style>
""", height=120)

st.markdown("""
<div style="background:#EAF3FF;border-left:4px solid #0B4F8A;
padding:12px 14px;border-radius:12px;font-weight:600;color:#0B4F8A;">
功能說明：上傳簡報 → 拆分任務 → 影片雲端化 → 內嵌優化 → Google Slides 發布 → 寫入資料庫
</div>
""", unsafe_allow_html=True)

# =========================================================
# 初始化狀態
# =========================================================
ensure_workspace()

if "split_jobs" not in st.session_state:
    st.session_state.split_jobs = []

if "ppt_meta" not in st.session_state:
    st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}

if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = None

if "bot" not in st.session_state:
    st.session_state.bot = PPTAutomationBot()

# =========================================================
# Step 1：選擇檔案
# =========================================================
st.subheader("步驟一：選擇檔案來源")
uploaded = st.file_uploader("PPTX", type=["pptx"])

source_path = os.path.join(WORK_DIR, "source.pptx")

if uploaded:
    filename = uploaded.name

    # 同檔名 → 保留拆分任務
    if st.session_state.current_file_name != filename:
        cleanup_workspace()
        st.session_state.split_jobs = load_history(filename)

    with open(source_path, "wb") as f:
        f.write(uploaded.getbuffer())

    st.session_state.current_file_name = filename

    # 解析簡報
    prs = Presentation(source_path)
    preview = []
    for i, slide in enumerate(prs.slides):
        title = slide.shapes.title.text if slide.shapes.title else "無標題"
        preview.append({"頁碼": i + 1, "內容摘要": title})

    st.session_state.ppt_meta = {
        "total_slides": len(prs.slides),
        "preview_data": preview
    }

    st.success(f"已讀取 {filename}（共 {len(prs.slides)} 頁）")

# =========================================================
# Step 2：拆分任務
# =========================================================
if st.session_state.current_file_name:
    st.subheader("步驟二：設定拆分任務")

    with st.expander("頁碼對照表"):
        st.dataframe(st.session_state.ppt_meta["preview_data"], use_container_width=True)

    if st.button("新增任務"):
        st.session_state.split_jobs.append({
            "id": str(uuid.uuid4()),
            "filename": "",
            "start": 1,
            "end": st.session_state.ppt_meta["total_slides"],
            "category": "清潔",
            "subcategory": "",
            "client": "",
            "keywords": ""
        })

    for i, job in enumerate(st.session_state.split_jobs):
        with st.container(border=True):
            c1, c2, c3 = st.columns([3, 1, 1])
            job["filename"] = c1.text_input("檔名", job["filename"], key=f"f{i}")
            job["start"] = c2.number_input("起始頁", 1, st.session_state.ppt_meta["total_slides"], job["start"], key=f"s{i}")
            job["end"] = c3.number_input("結束頁", 1, st.session_state.ppt_meta["total_slides"], job["end"], key=f"e{i}")

            m1, m2, m3, m4 = st.columns(4)
            job["category"] = m1.selectbox("類型", ["清潔", "配送", "購物", "AURO"], index=0, key=f"c{i}")
            job["subcategory"] = m2.text_input("子分類", job["subcategory"], key=f"sc{i}")
            job["client"] = m3.text_input("客戶", job["client"], key=f"cl{i}")
            job["keywords"] = m4.text_input("關鍵字", job["keywords"], key=f"k{i}")

    save_history(st.session_state.current_file_name, st.session_state.split_jobs)

# =========================================================
# Step 3：執行（含斷點續傳）
# =========================================================
if st.session_state.current_file_name and st.session_state.split_jobs:
    st.subheader("步驟三：開始執行")

    resume_step = detect_resume_step()
    st.info(f"偵測到可從步驟 {resume_step} 繼續執行")

    if st.button("執行自動化排程", use_container_width=True):
        bot = st.session_state.bot
        progress = st.progress(0)

        try:
            # Step 1
            if resume_step <= 1:
                progress.progress(10)
                bot.extract_and_upload_videos(source_path, os.path.join(WORK_DIR, "media"))

            # Step 2
            mod_path = os.path.join(WORK_DIR, "modified.pptx")
            if resume_step <= 2 or not os.path.exists(mod_path):
                progress.progress(30)
                bot.replace_videos_with_images(source_path, mod_path)

            # Step 3
            slim_path = os.path.join(WORK_DIR, "slim.pptx")
            if resume_step <= 3 or not os.path.exists(slim_path):
                progress.progress(50)
                bot.shrink_pptx(mod_path, slim_path)

            # Step 4
            progress.progress(70)
            results = bot.split_and_upload(
                slim_path,
                st.session_state.split_jobs,
                file_prefix=os.path.splitext(st.session_state.current_file_name)[0]
            )

            # Step 5
            progress.progress(90)
            final = bot.embed_videos_in_slides(results)

            bot.log_to_sheets(final)
            progress.progress(100)

            st.success("所有自動化流程執行完成")

            if st.button("返回並處理新檔"):
                reset_to_step1()

        except Exception as e:
            st.error(f"流程中斷：{e}")
            st.code(traceback.format_exc())
