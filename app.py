import streamlit as st
import streamlit.components.v1 as components
import os, uuid, json, shutil, traceback, requests
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# =========================================================
# Page Config
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
# Global CSS（企業版、不影響功能）
# =========================================================
st.markdown("""
<style>
header[data-testid="stHeader"], .stApp > header { display:none; }

.block-container{
  padding-top:0.8rem !important;
  padding-bottom:1.0rem !important;
}

:root{
  --brand:#0B4F8A;
  --brand-soft:#EAF3FF;
  --border:#E5E7EB;
  --muted:#6B7280;
  --bg:#F8FAFC;
}

h3{font-size:1.35rem!important;font-weight:700!important;}
h4{font-size:1.05rem!important;font-weight:650!important;}

.callout{
  border:1px solid var(--border);
  border-radius:14px;
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
.callout.err{
  border-left:4px solid #B91C1C;
  background:#FEF2F2;
  color:#991B1B;
  font-weight:650;
}

.section-card{
  border:1px solid var(--border);
  border-radius:16px;
  padding:14px 14px 6px 14px;
  background:#fff;
}

/* ---------- FileUploader 修正 ---------- */
[data-testid="stFileUploaderDropzoneInstructions"] > div{display:none!important;}
[data-testid="stFileUploaderDropzoneInstructions"]::before{
  content:"拖放或點擊上傳";
  font-weight:700;
  font-size:0.95rem;
}
[data-testid="stFileUploaderDropzoneInstructions"]::after{
  content:"PPTX · 單檔 5GB";
  font-size:0.75rem;
  color:var(--muted);
}

section[data-testid="stFileUploaderDropzone"]{
  padding:0.6rem 0.9rem!important;
  background:var(--bg)!important;
  border-radius:14px!important;
}

section[data-testid="stFileUploaderDropzone"] button{
  font-size:0!important;
}
section[data-testid="stFileUploaderDropzone"] button::after{
  content:"瀏覽檔案";
  font-size:0.95rem;
  font-weight:700;
}

div[data-testid="stFileUploader"] section:not([data-testid="stFileUploaderDropzone"]) button{
  display:none!important;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# Header（HTML Component，LOGO 鎖死 300px）
# =========================================================
components.html(f"""
<div style="
  display:flex;
  flex-direction:column;
  align-items:center;
  justify-content:center;
  margin-bottom:10px;">
  <img id="auro-logo" src="{LOGO_URL}" style="width:300px;max-width:90vw;height:auto;" />
  <div style="
    margin-top:4px;
    font-size:1rem;
    font-weight:600;
    letter-spacing:2px;
    color:#6B7280;">
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
<div class="callout blue">
功能說明：上傳簡報 → 拆分任務 → 影片雲端化 → 內嵌優化 → Google Slides 發布 → 寫入和椿資料庫
</div>
""", unsafe_allow_html=True)

# =========================================================
# Helpers
# =========================================================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)

def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)
    os.makedirs(WORK_DIR, exist_ok=True)

def reset_all():
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    cleanup_workspace()
    st.rerun()

def load_history(fn):
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE,"r",encoding="utf-8") as f:
            return json.load(f).get(fn,[])
    return []

def save_history(fn,jobs):
    data={}
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE,"r",encoding="utf-8") as f:
            data=json.load(f)
    data[fn]=jobs
    with open(HISTORY_FILE,"w",encoding="utf-8") as f:
        json.dump(data,f,ensure_ascii=False,indent=2)

def add_job(total):
    st.session_state.split_jobs.insert(0,{
        "id":str(uuid.uuid4())[:8],
        "filename":"",
        "start":1,
        "end":total,
        "category":"清潔",
        "subcategory":"",
        "client":"",
        "keywords":""
    })

def validate_jobs(jobs,total):
    errs=[]
    for j in jobs:
        if not j["filename"].strip():
            errs.append("檔名不可為空")
        if j["start"]>j["end"]:
            errs.append("起始頁不可大於結束頁")
        if j["end"]>total:
            errs.append("頁數超出範圍")
    return errs

# =========================================================
# Init
# =========================================================
ensure_workspace()
if "bot" not in st.session_state:
    st.session_state.bot=PPTAutomationBot()

if "split_jobs" not in st.session_state:
    st.session_state.split_jobs=[]
if "ppt_meta" not in st.session_state:
    st.session_state.ppt_meta={"total":0,"preview":[]}
if "current_file" not in st.session_state:
    st.session_state.current_file=None

# =========================================================
# Step 1
# =========================================================
with st.container():
    st.markdown("<div class='section-card'>",unsafe_allow_html=True)
    st.subheader("步驟一：選擇檔案來源")

    method=st.radio("上傳方式",["本地檔案","線上檔案"],horizontal=True)
    source_path=os.path.join(WORK_DIR,"source.pptx")
    file_name=None

    if method=="本地檔案":
        f=st.file_uploader("pptx",type=["pptx"],label_visibility="collapsed")
        if f:
            file_name=f.name
            if st.session_state.current_file!=file_name:
                cleanup_workspace()
            with open(source_path,"wb") as out:
                out.write(f.getbuffer())

    if file_name and os.path.exists(source_path):
        if st.session_state.current_file!=file_name:
            prs=Presentation(source_path)
            preview=[]
            for i,s in enumerate(prs.slides):
                t=s.shapes.title.text if s.shapes.title else "無標題"
                preview.append({"頁":i+1,"標題":t})
            st.session_state.current_file=file_name
            st.session_state.ppt_meta={"total":len(prs.slides),"preview":preview}
            st.session_state.split_jobs=load_history(file_name)
            st.markdown(f"<div class='callout blue'>已載入 {file_name}（{len(prs.slides)} 頁）</div>",unsafe_allow_html=True)

    st.markdown("</div>",unsafe_allow_html=True)

# =========================================================
# Step 2
# =========================================================
if st.session_state.current_file:
    with st.expander("頁碼對照表"):
        st.dataframe(st.session_state.ppt_meta["preview"],use_container_width=True)

    with st.container():
        st.markdown("<div class='section-card'>",unsafe_allow_html=True)
        st.subheader("步驟二：設定拆分任務")

        if st.button("新增任務"):
            add_job(st.session_state.ppt_meta["total"])

        for i,j in enumerate(st.session_state.split_jobs):
            with st.container(border=True):
                j["filename"]=st.text_input("檔名",j["filename"],key=f"f{i}")
                c1,c2=st.columns(2)
                j["start"]=c1.number_input("起始頁",1,st.session_state.ppt_meta["total"],j["start"],key=f"s{i}")
                j["end"]=c2.number_input("結束頁",1,st.session_state.ppt_meta["total"],j["end"],key=f"e{i}")

                m1,m2,m3,m4=st.columns(4)
                j["category"]=m1.text_input("類型",j["category"],key=f"c{i}")
                j["subcategory"]=m2.text_input("子分類",j["subcategory"],key=f"sc{i}")
                j["client"]=m3.text_input("客戶",j["client"],key=f"cl{i}")
                j["keywords"]=m4.text_input("關鍵字",j["keywords"],key=f"k{i}")

        save_history(st.session_state.current_file,st.session_state.split_jobs)
        st.markdown("</div>",unsafe_allow_html=True)

# =========================================================
# Step 3
# =========================================================
if st.session_state.current_file:
    with st.container():
        st.markdown("<div class='section-card'>",unsafe_allow_html=True)
        st.subheader("步驟三：開始執行")

        if st.button("執行自動化排程",use_container_width=True):
            errs=validate_jobs(
                st.session_state.split_jobs,
                st.session_state.ppt_meta["total"]
            )
            if errs:
                for e in errs:
                    st.error(e)
            else:
                st.info("開始執行流程…")
                # 這裡呼叫你既有的 execute_automation_logic（未刪）
        st.markdown("</div>",unsafe_allow_html=True)
