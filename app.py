import streamlit as st
import streamlit.components.v1 as components
import os
import uuid
import json
import shutil
import traceback
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# ==================================================
# Page Config
# ==================================================
st.set_page_config(
    page_title="Aurotek｜簡報案例自動化發布平台",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# ==================================================
# CSS（保留你目前穩定版本，只補必要 UI）
# ==================================================
st.markdown("""
<style>
header[data-testid="stHeader"]{display:none;}
.block-container{padding-top:.8rem;padding-bottom:1rem;}

:root{
 --blue:#0B4F8A;
 --blue-soft:#EAF3FF;
 --border:#E5E7EB;
 --text:#111827;
 --muted:#6B7280;
}

.auro-header{text-align:center;margin-bottom:8px;}
.auro-header img{width:300px;height:auto;}
.auro-sub{color:#6B7280;font-weight:600;letter-spacing:2px;}

.callout{
 border:1px solid var(--border);
 border-left:4px solid var(--blue);
 background:var(--blue-soft);
 padding:12px 14px;
 border-radius:14px;
 font-weight:600;
 color:var(--blue);
 margin:10px 0;
}

.section{
 border:1px solid var(--border);
 border-radius:16px;
 padding:14px;
 background:#fff;
 margin-bottom:16px;
}

.stProgress > div > div > div > div{
 font-weight:700;
 color:white;
}
</style>
""", unsafe_allow_html=True)

# ==================================================
# Helpers
# ==================================================
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
        with open(HISTORY_FILE,"r",encoding="utf-8") as f:
            return json.load(f).get(filename,[])
    except:
        return []

def save_history(filename, jobs):
    data={}
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE,"r",encoding="utf-8") as f:
                data=json.load(f)
        except:
            data={}
    data[filename]=jobs
    with open(HISTORY_FILE,"w",encoding="utf-8") as f:
        json.dump(data,f,ensure_ascii=False,indent=2)

# ==================================================
# Header
# ==================================================
st.markdown(f"""
<div class="auro-header">
  <img src="{LOGO_URL}" />
  <div class="auro-sub">簡報案例自動化發布平台</div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="callout">
功能說明：上傳簡報 → 影片雲端化 → 內嵌優化 → 拆分簡報 → Google Slides 發布 → 寫入資料庫
</div>
""", unsafe_allow_html=True)

# ==================================================
# Session Init
# ==================================================
if "bot" not in st.session_state:
    st.session_state.bot = PPTAutomationBot()

if "split_jobs" not in st.session_state:
    st.session_state.split_jobs = []

if "current_file" not in st.session_state:
    st.session_state.current_file = None

# ==================================================
# Step 1 Upload
# ==================================================
with st.container():
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.subheader("步驟一：選擇檔案")

    ensure_workspace()
    source_path = os.path.join(WORK_DIR,"source.pptx")

    uploaded = st.file_uploader("PPTX",type=["pptx"],label_visibility="collapsed")
    if uploaded:
        if st.session_state.current_file != uploaded.name:
            cleanup_workspace()
            st.session_state.split_jobs = load_history(uploaded.name)

        with open(source_path,"wb") as f:
            f.write(uploaded.getbuffer())

        prs = Presentation(source_path)
        total = len(prs.slides)
        st.session_state.current_file = uploaded.name

        st.markdown(
            f"<div class='callout'>已讀取 {uploaded.name}（共 {total} 頁）</div>",
            unsafe_allow_html=True
        )

    st.markdown('</div>', unsafe_allow_html=True)

# ==================================================
# Step 2 Split Jobs
# ==================================================
if st.session_state.current_file:
    with st.container():
        st.markdown('<div class="section">', unsafe_allow_html=True)
        st.subheader("步驟二：設定拆分任務")

        if st.button("新增任務"):
            st.session_state.split_jobs.append({
                "id":str(uuid.uuid4()),
                "filename":"",
                "start":1,
                "end":1
            })

        for i,job in enumerate(st.session_state.split_jobs):
            with st.container(border=True):
                c1,c2,c3=st.columns([3,1,1])
                job["filename"]=c1.text_input("檔名",job["filename"],key=f"f{i}")
                job["start"]=c2.number_input("起始頁",1,999,job["start"],key=f"s{i}")
                job["end"]=c3.number_input("結束頁",1,999,job["end"],key=f"e{i}")

        save_history(st.session_state.current_file, st.session_state.split_jobs)
        st.markdown('</div>', unsafe_allow_html=True)

# ==================================================
# Step 3 Execute（補回百分比）
# ==================================================
if st.session_state.current_file:
    with st.container():
        st.markdown('<div class="section">', unsafe_allow_html=True)
        st.subheader("步驟三：開始執行")

        if st.button("執行自動化排程",type="primary",use_container_width=True):

            bot = st.session_state.bot
            progress = st.progress(0,text="準備開始…")

            try:
                # STEP 1
                def cb1(cur,total):
                    pct=int(cur/total*100) if total else 0
                    progress.progress(int(pct*0.2),text=f"步驟 1/5：影片處理中（{pct}%）")

                video_map = bot.extract_and_upload_videos(
                    source_path,
                    os.path.join(WORK_DIR,"media"),
                    progress_callback=cb1
                )

                # STEP 2
                def cb2(cur,total):
                    pct=int(cur/total*100)
                    progress.progress(20+int(pct*0.15),text=f"步驟 2/5：影片置換（{pct}%）")

                mod = os.path.join(WORK_DIR,"mod.pptx")
                bot.replace_videos_with_images(source_path,mod,video_map,progress_callback=cb2)

                # STEP 3
                def cb3(cur,total):
                    pct=int(cur/total*100)
                    progress.progress(35+int(pct*0.15),text=f"步驟 3/5：檔案優化（{pct}%）")

                slim = os.path.join(WORK_DIR,"slim.pptx")
                bot.shrink_pptx(mod,slim,progress_callback=cb3)

                # STEP 4
                def cb4(fname,cur,total):
                    pct=int(cur/total*100)
                    progress.progress(50+int(pct*0.3),text=f"步驟 4/5：拆分上傳 {fname}（{pct}%）")

                results = bot.split_and_upload(
                    slim,
                    st.session_state.split_jobs,
                    file_prefix=os.path.splitext(st.session_state.current_file)[0],
                    progress_callback=cb4
                )

                # STEP 5
                def cb5(cur,total):
                    pct=int(cur/total*100)
                    progress.progress(80+int(pct*0.2),text=f"步驟 5/5：播放器優化（{pct}%）")

                final = bot.embed_videos_in_slides(results,progress_callback=cb5)
                bot.log_to_sheets(final)

                progress.progress(100,text="流程完成")

                # ===============================
                # 完成圖卡（含複製）
                # ===============================
                cards = ""
                for r in final:
                    cards+=f"""
                    <div class="card">
                      <b>{r['filename']}</b>
                      <div>
                        <a href="{r['final_link']}" target="_blank">開啟</a>
                        <button onclick="navigator.clipboard.writeText('{r['final_link']}')">複製連結</button>
                      </div>
                    </div>
                    """

                components.html(f"""
                <style>
                .card{{border:1px solid #E5E7EB;border-radius:12px;padding:10px;margin:8px 0}}
                </style>
                <h3>完成結果</h3>
                {cards}
                """,height=320,scrolling=True)

            except Exception as e:
                st.error(str(e))
                st.code(traceback.format_exc())

        st.markdown('</div>', unsafe_allow_html=True)
