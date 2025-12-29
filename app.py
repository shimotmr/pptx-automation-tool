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

# ==========================================
# 基本設定
# ==========================================
st.set_page_config(
    page_title="Aurotek｜簡報案例自動化發布平台",
    page_icon="📊",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# ==========================================
# CSS（企業版，省略說明，與你前版一致）
# ==========================================
st.markdown("""
<style>
header[data-testid="stHeader"] { display:none; }
.block-container { padding-top:0.8rem; }

.auro-header{
  display:flex;
  flex-direction:column;
  align-items:center;
}
.auro-header img{ width:300px; max-width:90vw; }
.auro-sub{ color:#6B7280; font-weight:600; letter-spacing:2px; }

.callout{
  border:1px solid #E5E7EB;
  border-left:4px solid #0B4F8A;
  background:#EAF3FF;
  padding:12px;
  border-radius:12px;
  font-weight:600;
}
.callout.err{
  border-left-color:#B91C1C;
  background:#FEF2F2;
  color:#991B1B;
}
.section{
  border:1px solid #E5E7EB;
  border-radius:16px;
  padding:16px;
  margin-bottom:16px;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# Helper：工作區
# ==========================================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)

def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)
    os.makedirs(WORK_DIR, exist_ok=True)

# ==========================================
# Helper：歷史任務（斷點續傳）
# ==========================================
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

# ==========================================
# ✅ 關鍵修正：安全處理影片替換
# ==========================================
def safe_replace_videos(
    bot,
    source_path,
    output_path,
    video_map,
):
    """
    - 有影片：正常 replace
    - 沒影片：直接 copy source → output
    """
    if video_map and isinstance(video_map, dict) and len(video_map) > 0:
        bot.replace_videos_with_images(
            source_path,
            output_path,
            video_map
        )
        return "replaced"
    else:
        shutil.copyfile(source_path, output_path)
        return "skipped"

# ==========================================
# Header
# ==========================================
st.markdown(f"""
<div class="auro-header">
  <img src="{LOGO_URL}">
  <div class="auro-sub">簡報案例自動化發布平台</div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="callout">
功能流程：上傳簡報 → 拆分任務 → 影片雲端化（如有） → 簡報優化 → Google Slides 發布 → 寫入資料庫
</div>
""", unsafe_allow_html=True)

# ==========================================
# 初始化 Session
# ==========================================
if "split_jobs" not in st.session_state:
    st.session_state.split_jobs = []
if "ppt_meta" not in st.session_state:
    st.session_state.ppt_meta = {}
if "current_file" not in st.session_state:
    st.session_state.current_file = None
if "bot" not in st.session_state:
    st.session_state.bot = PPTAutomationBot()

# ==========================================
# Step 1：上傳檔案
# ==========================================
with st.container():
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.subheader("步驟一：選擇檔案")

    ensure_workspace()
    source_path = os.path.join(WORK_DIR, "source.pptx")

    uploaded = st.file_uploader("PPTX", type=["pptx"], label_visibility="collapsed")
    if uploaded:
        if st.session_state.current_file != uploaded.name:
            cleanup_workspace()
            st.session_state.split_jobs = load_history(uploaded.name)

        with open(source_path, "wb") as f:
            f.write(uploaded.getbuffer())

        prs = Presentation(source_path)
        total = len(prs.slides)
        preview = []
        for i, s in enumerate(prs.slides):
            t = s.shapes.title.text if s.shapes.title else "無標題"
            preview.append({"頁碼": i + 1, "標題": t})

        st.session_state.current_file = uploaded.name
        st.session_state.ppt_meta = {
            "total": total,
            "preview": preview
        }

        st.markdown(
            f"<div class='callout'>已讀取 {uploaded.name}（共 {total} 頁）</div>",
            unsafe_allow_html=True
        )

    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# Step 2：拆分任務
# ==========================================
if st.session_state.current_file:
    with st.expander("頁碼對照表", expanded=False):
        st.dataframe(
            st.session_state.ppt_meta["preview"],
            use_container_width=True,
            hide_index=True
        )

    with st.container():
        st.markdown("<div class='section'>", unsafe_allow_html=True)
        st.subheader("步驟二：設定拆分任務")

        if st.button("新增任務"):
            st.session_state.split_jobs.append({
                "id": str(uuid.uuid4()),
                "filename": "",
                "start": 1,
                "end": st.session_state.ppt_meta["total"],
                "category": "清潔",
                "subcategory": "",
                "client": "",
                "keywords": ""
            })

        for i, job in enumerate(st.session_state.split_jobs):
            with st.container(border=True):
                st.markdown(f"**任務 {i+1}**")
                c1, c2, c3 = st.columns([3, 1.5, 1.5])
                job["filename"] = c1.text_input("檔名", job["filename"], key=f"f{i}")
                job["start"] = c2.number_input(
                    "起始頁", 1, st.session_state.ppt_meta["total"], job["start"], key=f"s{i}"
                )
                job["end"] = c3.number_input(
                    "結束頁", 1, st.session_state.ppt_meta["total"], job["end"], key=f"e{i}"
                )

                m1, m2, m3, m4 = st.columns(4)
                job["category"] = m1.text_input("類型", job["category"], key=f"c{i}")
                job["subcategory"] = m2.text_input("子分類", job["subcategory"], key=f"sc{i}")
                job["client"] = m3.text_input("客戶", job["client"], key=f"cl{i}")
                job["keywords"] = m4.text_input("關鍵字", job["keywords"], key=f"k{i}")

        save_history(st.session_state.current_file, st.session_state.split_jobs)
        st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# Step 3：執行
# ==========================================
if st.session_state.current_file:
    with st.container():
        st.markdown("<div class='section'>", unsafe_allow_html=True)
        st.subheader("步驟三：開始執行")

        if st.button("執行自動化排程", use_container_width=True):
            try:
                bot = st.session_state.bot

                # Step 1：影片雲端化（可能為空）
                video_map_path = os.path.join(WORK_DIR, "video_map.json")
                if os.path.exists(video_map_path):
                    with open(video_map_path, "r", encoding="utf-8") as f:
                        video_map = json.load(f)
                else:
                    video_map = bot.extract_and_upload_videos(
                        source_path,
                        os.path.join(WORK_DIR, "media")
                    )
                    with open(video_map_path, "w", encoding="utf-8") as f:
                        json.dump(video_map, f, indent=2)

                # Step 2：安全影片替換
                modified = os.path.join(WORK_DIR, "modified.pptx")
                result = safe_replace_videos(
                    bot,
                    source_path,
                    modified,
                    video_map
                )

                if result == "skipped":
                    st.markdown(
                        "<div class='callout'>未偵測到影片，已略過影片相關步驟</div>",
                        unsafe_allow_html=True
                    )

                # Step 3：瘦身
                slim = os.path.join(WORK_DIR, "slim.pptx")
                bot.shrink_pptx(modified, slim)

                # Step 4：拆分上傳
                results = bot.split_and_upload(
                    slim,
                    st.session_state.split_jobs,
                    file_prefix=os.path.splitext(st.session_state.current_file)[0]
                )

                # Step 5：內嵌影片（若有）
                final = bot.embed_videos_in_slides(results)

                bot.log_to_sheets(final)

                st.markdown(
                    "<div class='callout'>流程完成</div>",
                    unsafe_allow_html=True
                )

            except Exception as e:
                st.markdown(
                    f"<div class='callout err'>流程中斷：{e}</div>",
                    unsafe_allow_html=True
                )
                st.code(traceback.format_exc())

        st.markdown("</div>", unsafe_allow_html=True)
