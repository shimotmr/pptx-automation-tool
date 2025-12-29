import streamlit as st
import streamlit.components.v1 as components
import os, uuid, json, shutil, traceback, requests
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# ===============================
# 基本設定
# ===============================
st.set_page_config(
    page_title="Aurotek｜簡報案例自動化發布平台",
    page_icon="📊",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"
VIDEO_MAP_FILE = os.path.join(WORK_DIR, "video_map.json")

# ===============================
# CSS（企業藍、無綠色）
# ===============================
st.markdown("""
<style>
header[data-testid="stHeader"]{display:none;}
.block-container{padding-top:0.8rem;}

:root{
 --brand:#0B4F8A;
 --brand-bg:#EAF3FF;
 --border:#E5E7EB;
 --muted:#6B7280;
}

.callout{
 border-left:4px solid var(--brand);
 background:var(--brand-bg);
 border-radius:12px;
 padding:12px 14px;
 font-weight:650;
 color:var(--brand);
 margin:8px 0;
}
</style>
""", unsafe_allow_html=True)

# ===============================
# Helper
# ===============================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)

def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)
    os.makedirs(WORK_DIR, exist_ok=True)

def load_history(filename):
    if not os.path.exists(HISTORY_FILE):
        return []
    with open(HISTORY_FILE,"r",encoding="utf-8") as f:
        return json.load(f).get(filename,[])

def save_history(filename, jobs):
    data={}
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE,"r",encoding="utf-8") as f:
            data=json.load(f)
    data[filename]=jobs
    with open(HISTORY_FILE,"w",encoding="utf-8") as f:
        json.dump(data,f,ensure_ascii=False,indent=2)

# ===============================
# Header
# ===============================
components.html(f"""
<div style="text-align:center">
 <img src="{LOGO_URL}" style="width:300px;max-width:90vw"/>
 <div style="margin-top:4px;font-weight:600;letter-spacing:2px;color:#6B7280">
  簡報案例自動化發布平台
 </div>
</div>
<style>
@media (max-width:768px){{
 img{{width:260px !important;}}
}}
</style>
""", height=120)

st.markdown(
"<div class='callout'>功能說明：上傳簡報 → 拆分任務 → 影片雲端化 → 內嵌優化 → Google Slides 發布 → 寫入資料庫</div>",
unsafe_allow_html=True
)

# ===============================
# Init state
# ===============================
if "split_jobs" not in st.session_state:
    st.session_state.split_jobs=[]
if "ppt_meta" not in st.session_state:
    st.session_state.ppt_meta={"total":0}
if "current_file" not in st.session_state:
    st.session_state.current_file=None
if "bot" not in st.session_state:
    st.session_state.bot=PPTAutomationBot()

ensure_workspace()
source_path=os.path.join(WORK_DIR,"source.pptx")

# ===============================
# Step 1
# ===============================
st.subheader("步驟一：選擇檔案來源")
uploaded=st.file_uploader("PPTX",type=["pptx"],label_visibility="collapsed")

if uploaded:
    if st.session_state.current_file!=uploaded.name:
        cleanup_workspace()
        st.session_state.split_jobs=load_history(uploaded.name)

    with open(source_path,"wb") as f:
        f.write(uploaded.getbuffer())

    prs=Presentation(source_path)
    st.session_state.ppt_meta["total"]=len(prs.slides)
    st.session_state.current_file=uploaded.name

    st.markdown(
        f"<div class='callout'>已讀取 {uploaded.name}（共 {len(prs.slides)} 頁）</div>",
        unsafe_allow_html=True
    )

# ===============================
# Step 2（欄位完整）
# ===============================
if st.session_state.current_file:
    st.subheader("步驟二：設定拆分任務")

    if st.button("新增任務"):
        st.session_state.split_jobs.append({
            "id":str(uuid.uuid4())[:8],
            "filename":"",
            "start":1,
            "end":st.session_state.ppt_meta["total"],
            "category":"清潔",
            "subcategory":"",
            "client":"",
            "keywords":""
        })

    for i,job in enumerate(st.session_state.split_jobs):
        with st.container(border=True):
            job["filename"]=st.text_input("檔名",job["filename"],key=f"f{i}")
            c1,c2=st.columns(2)
            job["start"]=c1.number_input("起始頁",1,st.session_state.ppt_meta["total"],job["start"],key=f"s{i}")
            job["end"]=c2.number_input("結束頁",1,st.session_state.ppt_meta["total"],job["end"],key=f"e{i}")

            m1,m2,m3,m4=st.columns(4)
            job["category"]=m1.selectbox("類型",["清潔","配送","購物","AURO"],index=0,key=f"cat{i}")
            job["subcategory"]=m2.text_input("子分類",job["subcategory"],key=f"sub{i}")
            job["client"]=m3.text_input("客戶",job["client"],key=f"cli{i}")
            job["keywords"]=m4.text_input("關鍵字",job["keywords"],key=f"key{i}")

    save_history(st.session_state.current_file,st.session_state.split_jobs)

# ===============================
# Step 3（修正 video_map）
# ===============================
if st.session_state.current_file:
    st.subheader("步驟三：開始執行")

    if os.path.exists(VIDEO_MAP_FILE):
        st.markdown("<div class='callout'>偵測到可從上次中斷點繼續執行</div>",unsafe_allow_html=True)

    if st.button("執行自動化排程",use_container_width=True):
        try:
            bot=st.session_state.bot

            # Step 1：video_map
            if os.path.exists(VIDEO_MAP_FILE):
                with open(VIDEO_MAP_FILE,"r",encoding="utf-8") as f:
                    video_map=json.load(f)
            else:
                video_map=bot.extract_and_upload_videos(
                    source_path,
                    os.path.join(WORK_DIR,"media"),
                    file_prefix=os.path.splitext(st.session_state.current_file)[0]
                )
                with open(VIDEO_MAP_FILE,"w",encoding="utf-8") as f:
                    json.dump(video_map,f,ensure_ascii=False,indent=2)

            # Step 2
            mod_path=os.path.join(WORK_DIR,"modified.pptx")
            if not os.path.exists(mod_path):
                bot.replace_videos_with_images(source_path,mod_path,video_map)

            # Step 3
            slim_path=os.path.join(WORK_DIR,"slim.pptx")
            if not os.path.exists(slim_path):
                bot.shrink_pptx(mod_path,slim_path)

            # Step 4
            results=bot.split_and_upload(
                slim_path,
                st.session_state.split_jobs,
                file_prefix=os.path.splitext(st.session_state.current_file)[0]
            )

            # Step 5
            final_results=bot.embed_videos_in_slides(results)
            bot.log_to_sheets(final_results)

            st.markdown("<div class='callout'>流程已完成</div>",unsafe_allow_html=True)

        except Exception as e:
            st.markdown(
                f"<div class='callout' style='border-left-color:#B91C1C;color:#991B1B;background:#FEF2F2;'>流程中斷：{e}</div>",
                unsafe_allow_html=True
            )
            st.code(traceback.format_exc())
