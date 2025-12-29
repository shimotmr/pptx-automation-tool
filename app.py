import streamlit as st
import streamlit.components.v1 as components
import os, json, uuid, shutil, traceback
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# =============================
# Page Config
# =============================
st.set_page_config(
    page_title="Aurotek｜簡報案例自動化發布平台",
    layout="wide"
)

WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# =============================
# CSS（只影響外觀）
# =============================
st.markdown("""
<style>
header[data-testid="stHeader"]{display:none;}
.block-container{padding-top:0.8rem;}
.section-card{border:1px solid #E5E7EB;border-radius:14px;padding:14px;background:#fff;margin-bottom:16px;}
.callout{border-left:4px solid #0B4F8A;background:#EAF3FF;padding:12px;border-radius:10px;font-weight:600;}
.progress-text{font-weight:600;}
</style>
""", unsafe_allow_html=True)

# =============================
# Helpers
# =============================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)

def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)
    os.makedirs(WORK_DIR, exist_ok=True)

def load_history(filename):
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f).get(filename, [])
    return []

def save_history(filename, jobs):
    data = {}
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    data[filename] = jobs
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def validate_jobs(jobs, total_slides):
    errors = []
    for j in jobs:
        if not j["filename"].strip():
            errors.append("檔名不可為空")
        if j["start"] > j["end"]:
            errors.append("起始頁不可大於結束頁")
        if j["end"] > total_slides:
            errors.append("結束頁超出簡報總頁數")
    return errors

def safe_replace_videos(bot, source, target, video_map, progress_cb):
    if not video_map:
        progress_cb(1, 1)
        return source
    bot.replace_videos_with_images(
        source, target, video_map,
        progress_callback=progress_cb
    )
    return target

# =============================
# Session Init
# =============================
if "step1_done" not in st.session_state:
    st.session_state.step1_done = False
if "split_jobs" not in st.session_state:
    st.session_state.split_jobs = []
if "ppt_meta" not in st.session_state:
    st.session_state.ppt_meta = {}
if "current_file" not in st.session_state:
    st.session_state.current_file = None

bot = PPTAutomationBot()

# =============================
# Step 1：Upload
# =============================
with st.container():
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("步驟一：上傳簡報")

    ensure_workspace()
    source_path = os.path.join(WORK_DIR, "source.pptx")

    f = st.file_uploader("PPTX", type=["pptx"])
    if f:
        if st.session_state.current_file != f.name:
            cleanup_workspace()
            st.session_state.split_jobs = load_history(f.name)

        with open(source_path, "wb") as w:
            w.write(f.getbuffer())

        prs = Presentation(source_path)
        st.session_state.ppt_meta = {
            "total": len(prs.slides)
        }
        st.session_state.current_file = f.name
        st.session_state.step1_done = True

        st.markdown(f"<div class='callout'>已讀取 {f.name}（{len(prs.slides)} 頁）</div>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

# =============================
# Step 2：Tasks
# =============================
if st.session_state.step1_done:
    with st.container():
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("步驟二：設定拆分任務")

        if st.button("新增任務"):
            st.session_state.split_jobs.insert(0, {
                "id": str(uuid.uuid4()),
                "filename": "",
                "start": 1,
                "end": st.session_state.ppt_meta["total"],
                "category": "",
                "subcategory": "",
                "client": "",
                "keywords": ""
            })

        for idx, job in enumerate(st.session_state.split_jobs):
            st.markdown(f"**任務 {idx+1}**")
            c1, c2, c3 = st.columns([3,1,1])
            job["filename"] = c1.text_input("檔名", job["filename"], key=f"fn_{job['id']}")
            job["start"] = c2.number_input("起始頁", 1, st.session_state.ppt_meta["total"], job["start"], key=f"s_{job['id']}")
            job["end"] = c3.number_input("結束頁", 1, st.session_state.ppt_meta["total"], job["end"], key=f"e_{job['id']}")

            m1, m2, m3, m4 = st.columns(4)
            job["category"] = m1.text_input("類型", job["category"], key=f"cat_{job['id']}")
            job["subcategory"] = m2.text_input("子分類", job["subcategory"], key=f"sub_{job['id']}")
            job["client"] = m3.text_input("客戶", job["client"], key=f"cli_{job['id']}")
            job["keywords"] = m4.text_input("關鍵字", job["keywords"], key=f"key_{job['id']}")

            if st.button("刪除", key=f"d_{job['id']}"):
                st.session_state.split_jobs.pop(idx)
                st.experimental_rerun()

        save_history(st.session_state.current_file, st.session_state.split_jobs)
        st.markdown("</div>", unsafe_allow_html=True)

# =============================
# Step 3：Execute
# =============================
if st.session_state.step1_done and st.session_state.split_jobs:
    with st.container():
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("步驟三：開始執行")

        if st.button("執行自動化排程"):
            errors = validate_jobs(
                st.session_state.split_jobs,
                st.session_state.ppt_meta["total"]
            )
            if errors:
                for e in errors:
                    st.error(e)
                st.stop()

            progress = st.progress(0)
            text = st.empty()

            text.markdown("影片解析中 0%")
            video_map = bot.extract_and_upload_videos(
                source_path,
                os.path.join(WORK_DIR, "media"),
                progress_callback=lambda c,t: (
                    progress.progress(int(c/t*20)),
                    text.markdown(f"影片處理中 {int(c/t*100)}%")
                )
            )

            mod = os.path.join(WORK_DIR, "mod.pptx")
            safe_replace_videos(
                bot, source_path, mod, video_map,
                lambda c,t: (
                    progress.progress(20 + int(c/t*20)),
                    text.markdown(f"影片置換 {int(c/t*100)}%")
                )
            )

            slim = os.path.join(WORK_DIR, "slim.pptx")
            bot.shrink_pptx(
                mod, slim,
                progress_callback=lambda c,t: (
                    progress.progress(40 + int(c/t*20)),
                    text.markdown(f"檔案優化 {int(c/t*100)}%")
                )
            )

            results = bot.split_and_upload(
                slim,
                st.session_state.split_jobs,
                progress_callback=lambda c,t: (
                    progress.progress(60 + int(c/t*20)),
                    text.markdown(f"簡報上傳 {int(c/t*100)}%")
                )
            )

            final = bot.embed_videos_in_slides(
                results,
                progress_callback=lambda c,t: (
                    progress.progress(80 + int(c/t*20)),
                    text.markdown(f"播放器優化 {int(c/t*100)}%")
                )
            )

            progress.progress(100)
            text.markdown("完成 100%")

            st.markdown("<div class='callout'>流程完成</div>", unsafe_allow_html=True)

            for r in final:
                st.write(r["filename"], r["final_link"])

        st.markdown("</div>", unsafe_allow_html=True)
