# Version: v1.7 (Title Preview Fixed)
# Update Log:
# 1. FIXED: Restored logic to extract Slide Titles/Content for the preview table.
# 2. CORE: Kept v1.6.1 Batch Processing & GC to prevent OOM with 29+ videos.
# 3. UI: Maintained Blue styling, alignment, and clean headers.

import streamlit as st
import streamlit.components.v1 as components
import os
import uuid
import json
import shutil
import traceback
import requests
import gc
import math
from pptx import Presentation

# -------------------------------------------------
# 1. 依賴檢查
# -------------------------------------------------
try:
    from ppt_processor import PPTAutomationBot
except ImportError:
    st.error("❌ 嚴重錯誤：找不到 `ppt_processor.py`，請確認檔案已上傳至同一目錄。")
    st.stop()

# ==========================================
# 2. 設定與 CSS
# ==========================================
st.set_page_config(
    page_title="Aurotek 自動化發布平台",
    page_icon="📄",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

st.markdown("""
<style>
/* 隱藏 Header */
header[data-testid="stHeader"], .stApp > header { display: none; }
.block-container { padding-top: 1rem !important; padding-bottom: 6rem !important; }

/* 藍色風格提示 */
div[data-testid="stAlert"][data-style="success"], 
div[data-testid="stAlert"][data-style="info"] { 
    background-color: #F0F2F6 !important; 
    color: #31333F !important; 
    border: 1px solid #d0d7de !important; 
}
div[data-testid="stAlert"] svg { color: #004280 !important; }

/* 上傳按鈕優化 */
[data-testid="stFileUploaderDropzoneInstructions"] > div { display: none !important; }
[data-testid="stFileUploaderDropzoneInstructions"]::before {
    content: "請將檔案拖放至此"; display: block; font-weight: 700; color: #31333F; margin-bottom: 4px;
}
section[data-testid="stFileUploaderDropzone"] button {
    border: 1px solid #d0d7de; background: #fff; color: transparent !important;
    position: relative; border-radius: 4px; min-height: 38px; width: auto; margin-top: 10px;
}
section[data-testid="stFileUploaderDropzone"] button::after {
    content: "瀏覽檔案"; position: absolute; color: #31333F; left: 50%; top: 50%; transform: translate(-50%, -50%);
    white-space: nowrap; font-weight: 500; font-size: 14px;
}
/* 排除刪除按鈕 */
[data-testid="stFileUploaderDeleteBtn"] { border: none !important; background: transparent; color: inherit !important; }
[data-testid="stFileUploaderDeleteBtn"]::after { content: none; }

/* 垃圾桶與按鈕對齊 */
div[data-testid="column"] button { 
    border: 1px solid #eee !important; background: white !important; color: #555 !important; 
    font-size: 0.85rem !important; min-width: 40px !important; padding: 4px 8px !important; 
}
div[data-testid="column"] button:hover { 
    color: #cc0000 !important; border-color: #cc0000 !important; background: #fff5f5 !important; 
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. 核心功能函數
# ==========================================
def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        try: shutil.rmtree(WORK_DIR)
        except: pass
    os.makedirs(WORK_DIR, exist_ok=True)

def reset_callback():
    cleanup_workspace()
    if st.session_state.get('current_file_name') and os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f: data = json.load(f)
            if st.session_state.current_file_name in data:
                del data[st.session_state.current_file_name]
                with open(HISTORY_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
        except: pass
        
    st.session_state.split_jobs = []
    st.session_state.current_file_name = None
    st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}
    st.session_state.execution_results = None 
    st.session_state.reset_key += 1
    gc.collect()

def load_history(filename):
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f: return json.load(f).get(filename, [])
        except: return []
    return []

def save_history(filename, jobs):
    try:
        data = {}
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, "r", encoding="utf-8") as f: 
                try: data = json.load(f)
                except: pass
        data[filename] = jobs
        with open(HISTORY_FILE, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
    except: pass

def add_split_job(total_pages):
    st.session_state.split_jobs.insert(0, {
        "id": str(uuid.uuid4())[:8], "filename": "", "start": 1, "end": total_pages,
        "category": "清潔", "subcategory": "", "client": "", "keywords": ""
    })

def remove_split_job(index):
    st.session_state.split_jobs.pop(index)

def validate_jobs(jobs, total_slides):
    errors = []
    for i, job in enumerate(jobs):
        if not job['filename'].strip(): errors.append(f"❌ 任務 {len(jobs)-i}: 檔名為空")
        if job['start'] > job['end']: errors.append(f"❌ 任務 {len(jobs)-i}: 起始頁大於結束頁")
        if job['end'] > total_slides: errors.append(f"❌ 任務 {len(jobs)-i}: 結束頁超出範圍")
    return errors

def download_file_from_url(url, dest_path):
    try:
        response = requests.get(url, stream=True, timeout=60)
        response.raise_for_status()
        with open(dest_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192): f.write(chunk)
        return True, None
    except Exception as e: return False, str(e)

def scroll_to_step4():
    components.html("""<script>setTimeout(function(){try{const s=window.parent.document.getElementById('step4-anchor');if(s)s.scrollIntoView({behavior:'smooth',block:'start'});}catch(e){}},500);</script>""", height=0)

def render_copy_btn(text):
    return f"""<html><body style="margin:0;padding:0;"><button onclick="navigator.clipboard.writeText('{text}')" style="border:1px solid #004280;background:#fff;color:#004280;padding:4px 8px;border-radius:4px;cursor:pointer;font-size:13px;">📋 複製</button></body></html>"""

# ==========================================
# 4. 核心執行邏輯 (分批處理 + 記憶體優化)
# ==========================================
def execute_automation_logic(bot, source_path, file_prefix, jobs, auto_clean):
    main_progress = st.progress(0, text="準備開始...")
    status_area = st.empty()
    detail_bar = st.empty()
    
    def update_bar(text, pct):
        detail_bar.progress(pct, text=text)

    try:
        # Step 1
        status_area.info("1️⃣ 步驟 1/5：提取 PPT 內影片並上傳至雲端...")
        main_progress.progress(5, text="Step 1: 影片雲端化")
        
        video_map = bot.extract_and_upload_videos(
            source_path,
            os.path.join(WORK_DIR, "media"),
            file_prefix=file_prefix,
            progress_callback=lambda f, c, t: update_bar(f"上傳中: {f}", c/t if t else 0),
            log_callback=print
        )
        gc.collect()

        # Step 2: Batch Processing
        status_area.info("2️⃣ 步驟 2/5：置換影片連結 (分批處理模式)...")
        main_progress.progress(25, text="Step 2: 連結置換")
        
        final_mod_path = os.path.join(WORK_DIR, "modified.pptx")
        temp_working_path = os.path.join(WORK_DIR, "temp_working.pptx")
        shutil.copy(source_path, temp_working_path)
        
        video_items = list(video_map.items())
        BATCH_SIZE = 5
        total_batches = math.ceil(len(video_items) / BATCH_SIZE)
        
        for i in range(0, len(video_items), BATCH_SIZE):
            batch_num = (i // BATCH_SIZE) + 1
            batch_items = dict(video_items[i : i + BATCH_SIZE])
            
            current_pct = batch_num / total_batches
            update_bar(f"批次處理 ({batch_num}/{total_batches}): 置換影片...", current_pct)
            
            temp_output = os.path.join(WORK_DIR, f"temp_batch_{batch_num}.pptx")
            
            bot.replace_videos_with_images(
                temp_working_path,
                temp_output,
                batch_items,
                progress_callback=None
            )
            
            if os.path.exists(temp_working_path): os.remove(temp_working_path)
            shutil.move(temp_output, temp_working_path)
            gc.collect() # 關鍵釋放
        
        if os.path.exists(final_mod_path): os.remove(final_mod_path)
        shutil.move(temp_working_path, final_mod_path)
        detail_bar.empty()

        # Step 3
        status_area.info("3️⃣ 步驟 3/5：進行檔案壓縮與瘦身...")
        main_progress.progress(45, text="Step 3: 檔案瘦身")
        slim_path = os.path.join(WORK_DIR, "slim.pptx")
        bot.shrink_pptx(final_mod_path, slim_path, progress_callback=lambda c, t: update_bar("壓縮中...", c/t if t else 0))
        gc.collect()

        # Step 4
        status_area.info("4️⃣ 步驟 4/5：依設定拆分簡報並上傳...")
        main_progress.progress(65, text="Step 4: 拆分發布")
        results = bot.split_and_upload(
            slim_path, sorted(jobs, key=lambda x: x['start']), file_prefix,
            progress_callback=lambda f, c, t: update_bar(f"上傳簡報: {f}", c/t if t else 0),
            log_callback=print
        )
        
        if any(r.get('error_too_large') for r in results):
            st.error("⛔️ 流程終止：部分檔案過大無法上傳。")
            return

        # Step 5
        status_area.info("5️⃣ 步驟 5/5：優化線上播放器...")
        main_progress.progress(85, text="Step 5: 內嵌優化")
        final_results = bot.embed_videos_in_slides(results, progress_callback=lambda c, t: update_bar("優化中...", c/t if t else 0), log_callback=print)

        # Final
        status_area.info("📝 最後步驟：寫入資料庫...")
        main_progress.progress(95, text="Final: 寫入資料庫")
        bot.log_to_sheets(final_results, log_callback=print)

        main_progress.progress(100, text="任務完成")
        status_area.info("**成功：** 所有自動化流程執行完畢。", icon=None)
        
        if auto_clean: cleanup_workspace()
        
        st.session_state.execution_results = {"results": final_results, "prefix": file_prefix}

    except Exception as e:
        st.error(f"❌ 執行流程發生錯誤: {str(e)}")
        with st.expander("查看詳細錯誤資訊"):
            st.code(traceback.format_exc())

# ==========================================
# 5. 主介面邏輯
# ==========================================
os.makedirs(WORK_DIR, exist_ok=True)

# Header
components.html(f"""<div style="width:100%;display:flex;flex-direction:column;align-items:center;margin:4px 0 2px 0;"><img src="{LOGO_URL}" style="width:300px;"><div style="margin-top:4px;color:gray;font-size:1rem;letter-spacing:2px;">簡報案例自動化發布平台</div></div>""", height=78)
st.info("功能說明： 上傳PPT → 線上拆分 → 影片雲端化 → 內嵌優化 → 簡報雲端化 → 寫入和椿資料庫")

# Init State
if 'split_jobs' not in st.session_state: st.session_state.split_jobs = []
if 'reset_key' not in st.session_state: st.session_state.reset_key = 0
if 'execution_results' not in st.session_state: st.session_state.execution_results = None
if 'bot' not in st.session_state:
    try:
        bot_instance = PPTAutomationBot()
        if bot_instance.creds: st.session_state.bot = bot_instance
    except: pass
if 'current_file_name' not in st.session_state: st.session_state.current_file_name = None
if 'ppt_meta' not in st.session_state: st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}

# Step 1
with st.container(border=True):
    st.subheader("步驟一：選擇檔案來源")
    input_method = st.radio("上傳方式", ["本地檔案", "線上檔案"], horizontal=True)
    uploaded_file = None
    source_path = os.path.join(WORK_DIR, "source.pptx")
    file_name_for_logic = None
    
    if input_method == "本地檔案":
        uploaded_file = st.file_uploader("請選擇 PPTX 檔案", type=['pptx'], label_visibility="collapsed", key=f"uploader_{st.session_state.reset_key}")
        if uploaded_file:
            file_name_for_logic = uploaded_file.name
            if st.session_state.current_file_name != file_name_for_logic:
                cleanup_workspace()
                with open(source_path, "wb") as f: f.write(uploaded_file.getbuffer())
            elif not os.path.exists(source_path):
                 with open(source_path, "wb") as f: f.write(uploaded_file.getbuffer())
    else:
        url_input = st.text_input("請輸入 PPTX 網址", key=f"url_{st.session_state.reset_key}")
        if url_input and st.button("下載"):
            cleanup_workspace()
            success, err = download_file_from_url(url_input, source_path)
            if success:
                file_name_for_logic = "downloaded.pptx"
                st.info("下載成功", icon="✅")
            else: st.error(f"下載失敗: {err}")

    if file_name_for_logic and os.path.exists(source_path):
        file_prefix = os.path.splitext(file_name_for_logic)[0]
        if st.session_state.current_file_name != file_name_for_logic:
            saved_jobs = load_history(file_name_for_logic)
            st.session_state.split_jobs = saved_jobs if saved_jobs else []
            try:
                prs = Presentation(source_path)
                total_slides = len(prs.slides)
                
                # [FIXED] 恢復標題讀取功能
                preview_data = []
                for i, slide in enumerate(prs.slides):
                    txt = slide.shapes.title.text if (slide.shapes.title and slide.shapes.title.text) else "無標題"
                    if txt == "無標題":
                        for s in slide.shapes:
                            if hasattr(s, "text") and s.text.strip():
                                txt = s.text.strip()[:20] + "..."
                                break
                    preview_data.append({"頁碼": i+1, "內容摘要": txt})
                
                st.session_state.ppt_meta["total_slides"] = total_slides
                st.session_state.ppt_meta["preview_data"] = preview_data
                st.session_state.current_file_name = file_name_for_logic
                st.session_state.execution_results = None 
                st.info(f"**已讀取：** {file_name_for_logic} (共 {len(prs.slides)} 頁)", icon=None)
            except Exception as e:
                st.error(f"檔案處理失敗: {e}")
                st.session_state.current_file_name = None
                st.stop()

# Step 2
if st.session_state.current_file_name:
    with st.expander("👁️ 查看頁碼對照表"):
        st.dataframe(st.session_state.ppt_meta["preview_data"], use_container_width=True)

    with st.container(border=True):
        c1, c2 = st.columns([3, 1])
        c1.subheader("步驟二：設定拆分任務")
        if c2.button("➕ 新增任務", type="primary", use_container_width=True):
            add_split_job(st.session_state.ppt_meta["total_slides"])

        if not st.session_state.split_jobs:
            st.info("尚未建立任務，請點擊上方按鈕新增。")
        
        for i, job in enumerate(st.session_state.split_jobs):
            with st.container(border=True):
                c_title, c_del = st.columns([0.85, 0.15])
                c_title.markdown(f"**任務 {len(st.session_state.split_jobs)-i}**")
                if c_del.button("🗑️ 刪除", key=f"del_{job['id']}"):
                    remove_split_job(i)
                    st.rerun()
                
                c_a, c_b, c_c = st.columns([3, 1.5, 1.5])
                job["filename"] = c_a.text_input("檔名", value=job["filename"], key=f"f_{job['id']}")
                job["start"] = c_b.number_input("起始", 1, st.session_state.ppt_meta["total_slides"], job["start"], key=f"s_{job['id']}")
                job["end"] = c_c.number_input("結束", 1, st.session_state.ppt_meta["total_slides"], job["end"], key=f"e_{job['id']}")
                
                c_d, c_e, c_f, c_g = st.columns(4)
                job["category"] = c_d.selectbox("類型", ["清潔", "配送", "購物", "AURO"], key=f"cat_{job['id']}")
                job["subcategory"] = c_e.text_input("子分類", value=job["subcategory"], key=f"sub_{job['id']}")
                job["client"] = c_f.text_input("客戶", value=job["client"], key=f"cli_{job['id']}")
                job["keywords"] = c_g.text_input("關鍵字", value=job["keywords"], key=f"key_{job['id']}")
        
        save_history(st.session_state.current_file_name, st.session_state.split_jobs)

    # Step 3
    if st.session_state.split_jobs:
        with st.container(border=True):
            st.subheader("步驟三：執行任務")
            auto_clean = st.checkbox("任務完成後自動清除暫存檔", value=True)
            if st.button("執行雲端化任務", type="primary", use_container_width=True):
                errs = validate_jobs(st.session_state.split_jobs, st.session_state.ppt_meta["total_slides"])
                if errs:
                    for e in errs: st.error(e)
                else:
                    if st.session_state.bot:
                        execute_automation_logic(
                            st.session_state.bot,
                            os.path.join(WORK_DIR, "source.pptx"),
                            os.path.splitext(st.session_state.current_file_name)[0],
                            st.session_state.split_jobs,
                            auto_clean
                        )
                        st.rerun()
                    else: st.error("Bot 未初始化")

# Step 4
if st.session_state.execution_results:
    st.markdown("<div id='step4-anchor'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.subheader("步驟四：產出結果")
        results = st.session_state.execution_results["results"]
        pfx = st.session_state.execution_results["prefix"]
        
        rows = ""
        for r in results:
            if 'final_link' in r:
                rows += f"""<tr style="border-bottom:1px solid #eee;"><td style="padding:8px;color:#333;">[{pfx}]_{r['filename']}</td><td style="padding:8px;"><a href="{r['final_link']}" target="_blank" style="text-decoration:none;color:#004280;font-weight:500;border:1px solid #004280;padding:4px 8px;border-radius:4px;display:inline-block;">開啟簡報</a></td><td style="padding:8px;">{render_copy_btn(r['final_link'])}</td></tr>"""
        
        if rows: st.markdown(f"""<table style="width:100%;font-size:14px;border-collapse:collapse;"><tr style="background-color:#f9f9f9;text-align:left;border-bottom:1px solid #ddd;"><th style="padding:8px;">檔案名稱</th><th style="padding:8px;">線上預覽</th><th style="padding:8px;">操作</th></tr>{rows}</table>""", unsafe_allow_html=True)
        else: st.warning("沒有產生任何結果。")
    scroll_to_step4()

# Footer
if st.session_state.current_file_name:
    st.markdown("<div style='margin-top: 40px;'></div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.button("清除任務，上傳新簡報", type="primary", on_click=reset_callback, use_container_width=True)
    with c2:
        st.link_button("前往「和椿數位資源庫」", "https://aurotek.pse.is/puducases", type="primary", use_container_width=True)