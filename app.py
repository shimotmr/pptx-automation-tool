import streamlit as st
import os
import uuid
import json
import shutil
import traceback
import requests
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# ==========================================
#              設定頁面與樣式
# ==========================================
st.set_page_config(
    page_title="Aurotek數位資料庫 簡報案例自動化發布平台",
    page_icon="🤖",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# ==========================================
#              CSS 深度優化
# ==========================================
st.markdown("""
    <style>
    /* 1. 隱藏 Streamlit 預設 Header 與 Toolbar */
    header[data-testid="stHeader"] {
        display: none;
    }
    .stApp > header {
        display: none;
    }
    
    /* 2. 調整頂部間距 */
    .block-container {
        padding-top: 2rem !important; 
    }

    /* 3. 上傳元件中文化 */
    [data-testid="stFileUploaderDropzoneInstructions"] > div:first-child { visibility: hidden; height: 0; }
    [data-testid="stFileUploaderDropzoneInstructions"] > div:nth-child(2) { visibility: hidden; height: 0; }
    [data-testid="stFileUploaderDropzoneInstructions"]::before {
        content: "請將檔案拖放至此";
        visibility: visible;
        display: block;
        font-size: 1.2rem;
        font-weight: bold;
        margin-bottom: 5px;
    }
    [data-testid="stFileUploaderDropzoneInstructions"]::after {
        content: "單一檔案限制 5GB • PPTX";
        visibility: visible;
        display: block;
        font-size: 0.8rem;
        color: gray;
    }
    [data-testid="stFileUploader"] button { color: transparent !important; position: relative; }
    [data-testid="stFileUploader"] button::after {
        content: "瀏覽檔案";
        color: #31333F;
        position: absolute;
        left: 50%; top: 50%; transform: translate(-50%, -50%);
        font-weight: 500;
    }

    /* 4. 通用樣式與進度條 */
    h3 { font-size: 1.5rem !important; font-weight: 600 !important; }
    h4 { font-size: 1.2rem !important; font-weight: 600 !important; color: #555; }
    .stProgress > div > div > div > div { color: white; font-weight: 500; }
    
    /* [新增] 縮小 st.info 功能說明區塊的文字大小 */
    [data-testid="stAlert"] p {
        font-size: 0.85rem !important; /* 縮小字體 */
        line-height: 1.4 !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
#              Helper Functions
# ==========================================
def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        try:
            shutil.rmtree(WORK_DIR)
        except Exception as e:
            print(f"Cleanup warning: {e}")
    os.makedirs(WORK_DIR, exist_ok=True)

def load_history(filename):
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                hist = json.load(f)
                return hist.get(filename, [])
        except:
            return []
    return []

def save_history(filename, jobs):
    try:
        data = {}
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                try:
                    data = json.load(f)
                except:
                    data = {}
        data[filename] = jobs
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"History save failed: {e}")

def add_split_job(total_pages):
    st.session_state.split_jobs.insert(0, {
        "id": str(uuid.uuid4())[:8],
        "filename": "",
        "start": 1,
        "end": total_pages,
        "category": "清潔",
        "subcategory": "",
        "client": "",
        "keywords": ""
    })

def remove_split_job(index):
    st.session_state.split_jobs.pop(index)

def validate_jobs(jobs, total_slides):
    errors = []
    for i, job in enumerate(jobs):
        task_label = f"任務 {i+1} (檔名: {job['filename'] or '未命名'})"
        if not job['filename'].strip():
            errors.append(f"❌ {task_label}: 檔案名稱不能為空。")
        if job['start'] > job['end']:
            errors.append(f"❌ {task_label}: 起始頁 ({job['start']}) 不能大於 結束頁 ({job['end']})。")
        if job['end'] > total_slides:
            errors.append(f"❌ {task_label}: 結束頁 ({job['end']}) 超出了簡報總頁數 ({total_slides})。")

    sorted_jobs = sorted(jobs, key=lambda x: x['start'])
    for i in range(len(sorted_jobs) - 1):
        current_job = sorted_jobs[i]
        next_job = sorted_jobs[i+1]
        
        if current_job['end'] >= next_job['start']:
            conflict_msg = (
                f"⚠️ 發現頁數重疊！\n"
                f"   - {current_job['filename']} (範圍 {current_job['start']}-{current_job['end']})\n"
                f"   - {next_job['filename']} (範圍 {next_job['start']}-{next_job['end']})\n"
                f"   請確認是否重複包含了第 {next_job['start']} 到 {current_job['end']} 頁。"
            )
            errors.append(conflict_msg)

    return errors

def download_file_from_url(url, dest_path):
    try:
        response = requests.get(url, stream=True, timeout=60)
        response.raise_for_status()
        with open(dest_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        return True, None
    except Exception as e:
        return False, str(e)

# ==========================================
#              Core Logic Function
# ==========================================
def execute_automation_logic(bot, source_path, file_prefix, jobs, auto_clean):
    main_progress = st.progress(0, text="準備開始...")
    status_area = st.empty() 
    detail_bar_placeholder = st.empty()

    sorted_jobs = sorted(jobs, key=lambda x: x['start'])
    
    def update_step1(filename, current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"Step 1 詳細進度: 正在上傳 `{filename}` ({int(pct*100)}%)")

    def update_step2(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"Step 2 詳細進度: 處理投影片 {current}/{total} ({int(pct*100)}%)")

    def update_step3(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"Step 3 詳細進度: 處理內部檔案 {current}/{total} ({int(pct*100)}%)")

    def update_step4(filename, current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"Step 4 詳細進度: 正在上傳 `{filename}` ({int(pct*100)}%)")

    def update_step5(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"Step 5 詳細進度: 優化任務 {current}/{total} ({int(pct*100)}%)")
    
    def general_log(msg):
        print(f"[Log] {msg}")

    try:
        status_area.info("1️⃣ 步驟 1/5：提取 PPT 內影片並上傳至雲端...")
        main_progress.progress(5, text="Step 1: 影片雲端化")
        video_map = bot.extract_and_upload_videos(
            source_path, 
            os.path.join(WORK_DIR, "media"), 
            file_prefix=file_prefix,
            progress_callback=update_step1,
            log_callback=general_log
        )
        detail_bar_placeholder.empty()
        
        status_area.info("2️⃣ 步驟 2/5：將 PPT 內的影片替換為雲端連結圖片...")
        main_progress.progress(25, text="Step 2: 連結置換")
        mod_path = os.path.join(WORK_DIR, "modified.pptx")
        bot.replace_videos_with_images(
            source_path, 
            mod_path, 
            video_map,
            progress_callback=update_step2
        )
        detail_bar_placeholder.empty()
        
        status_area.info("3️⃣ 步驟 3/5：進行檔案壓縮與瘦身 (提升解析度)...")
        main_progress.progress(45, text="Step 3: 檔案瘦身")
        slim_path = os.path.join(WORK_DIR, "slim.pptx")
        bot.shrink_pptx(
            mod_path, 
            slim_path,
            progress_callback=update_step3
        )
        detail_bar_placeholder.empty()
        
        status_area.info("4️⃣ 步驟 4/5：依設定拆分簡報並上傳至 Google Slides...")
        main_progress.progress(65, text="Step 4: 拆分發布")
        results = bot.split_and_upload(
            slim_path, 
            sorted_jobs,
            file_prefix=file_prefix,
            progress_callback=update_step4,
            log_callback=general_log
        )
        detail_bar_placeholder.empty()
        
        oversized_errors = [r for r in results if r.get('error_too_large')]
        if oversized_errors:
            st.error("⛔️ 流程終止：偵測到拆分後的檔案過大。")
            for err_job in oversized_errors:
                st.error(f"❌ 任務「{err_job['filename']}」壓縮後仍有 {err_job['size_mb']:.2f} MB，超過 Google 限制 (100MB)。")
            st.warning("💡 建議做法：請減少該任務的頁數範圍，將其拆分為多個小任務後重新執行。")
            return
        
        status_area.info("5️⃣ 步驟 5/5：優化線上簡報的影片播放器...")
        main_progress.progress(85, text="Step 5: 內嵌優化")
        final_results = bot.embed_videos_in_slides(
            results,
            progress_callback=update_step5,
            log_callback=general_log
        )
        detail_bar_placeholder.empty()
        
        status_area.info("📝 最後步驟：將成果寫入 Google Sheets 資料庫...")
        main_progress.progress(95, text="Final: 寫入資料庫")
        bot.log_to_sheets(
            final_results,
            log_callback=general_log
        )
        
        main_progress.progress(100, text="🎉 任務全部完成！")
        status_area.success("🎉 所有自動化流程執行完畢！")
        st.balloons()
        
        if auto_clean:
            cleanup_workspace()
            st.toast("已自動清除暫存檔案。", icon="🧹")
        
        st.divider()
        st.subheader("✅ 產出結果連結")
        result_count = 0
        for res in final_results:
            if 'final_link' in res:
                result_count += 1
                display_name = f"[{file_prefix}]_{res['filename']}"
                st.markdown(f"👉 **{display_name}**: [點擊開啟 Google Slides]({res['final_link']})")
        
        if result_count == 0:
            st.warning("沒有產生任何結果，請檢查是否有任務被跳過。")

    except Exception as e:
        st.error(f"執行過程中發生錯誤: {e}")
        with st.expander("查看詳細錯誤資訊"):
            st.code(traceback.format_exc())

# ==========================================
#              Main UI (Layout)
# ==========================================

# 0. 確保工作目錄存在（避免首次使用就寫檔失敗）
os.makedirs(WORK_DIR, exist_ok=True)

# 1. Header：用 st.image + 置中容器（width 一定生效）
c1, c2, c3 = st.columns([1, 1, 1])
with c2:
    st.image(LOGO_URL, width=150)  # ✅ 你要縮 50%：原 300 -> 150
    st.markdown(
        "<div style='text-align:center; color:gray; font-size:1.0rem; font-weight:500; letter-spacing:2px; margin-top:6px;'>"
        "簡報案例自動化發布平台"
        "</div>",
        unsafe_allow_html=True
    )


# 2. 功能說明 (已透過 CSS 將文字縮小)
st.info("功能說明： 上傳PPT → 線上拆分 → 影片雲端化 → 內嵌優化 → 簡報雲端化 → 寫入和椿資料庫")

# 3. 初始化
if 'split_jobs' not in st.session_state:
    st.session_state.split_jobs = []

if 'bot' not in st.session_state:
    try:
        bot_instance = PPTAutomationBot()
        if bot_instance.creds:
            st.session_state.bot = bot_instance
        else:
            st.warning("⚠️ 系統未檢測到有效憑證 (Secrets)。")
    except Exception as e:
        st.error(f"Bot 初始化失敗: {e}")

if 'current_file_name' not in st.session_state:
    st.session_state.current_file_name = None
if 'ppt_meta' not in st.session_state:
    st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}

# 4. 檔案來源選擇區塊
with st.container(border=True):
    st.subheader("📂 步驟一：選擇檔案來源")
    
    input_method = st.radio("請選擇上傳方式：", ["本地檔案上傳", "線上 PPT 網址"], horizontal=True)
    
    uploaded_file = None
    source_path = os.path.join(WORK_DIR, "source.pptx")
    file_name_for_logic = None 

    if input_method == "本地檔案上傳":
        uploaded_file = st.file_uploader("請選擇 PPTX 檔案", type=['pptx'], label_visibility="collapsed")
        if uploaded_file:
            file_name_for_logic = uploaded_file.name
            if not os.path.exists(WORK_DIR):
                os.makedirs(WORK_DIR, exist_ok=True)
            with open(source_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

    else:
        url_input = st.text_input("請輸入 PPTX 檔案的直接下載網址 (Direct URL)", placeholder="https://example.com/file.pptx")
        if url_input:
            if not url_input.lower().endswith(".pptx"):
                st.warning("⚠️ 網址結尾似乎不是 .pptx，請確認網址正確性。")
            
            fake_name = url_input.split("/")[-1].split("?")[0]
            if not fake_name.lower().endswith(".pptx"):
                fake_name += ".pptx"
            
            if st.button("📥 下載並處理此網址"):
                with st.spinner("正在從網址下載檔案..."):
                    if not os.path.exists(WORK_DIR):
                        os.makedirs(WORK_DIR, exist_ok=True)
                    success, error = download_file_from_url(url_input, source_path)
                    if success:
                        file_name_for_logic = fake_name
                        st.success("下載成功！")
                    else:
                        st.error(f"下載失敗: {error}")

    # 5. 檔案處理邏輯
    if file_name_for_logic and os.path.exists(source_path):
        file_prefix = os.path.splitext(file_name_for_logic)[0]
        
        if st.session_state.current_file_name != file_name_for_logic:
            cleanup_workspace() 
            saved_jobs = load_history(file_name_for_logic)
            st.session_state.split_jobs = saved_jobs if saved_jobs else []
            
            progress_placeholder = st.empty()
            progress_placeholder.progress(0, text="解析檔案中...")
            
            try:
                prs = Presentation(source_path)
                total_slides = len(prs.slides)
                
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
                
                progress_placeholder.progress(100, text="完成！")
                st.success(f"✅ 已讀取：{file_name_for_logic} (共 {total_slides} 頁)")
                
            except Exception as e:
                st.error(f"檔案處理失敗: {e}")
                st.session_state.current_file_name = None
                st.stop()

if st.session_state.current_file_name:
    total_slides = st.session_state.ppt_meta["total_slides"]
    preview_data = st.session_state.ppt_meta["preview_data"]

    with st.expander("👁️ 點擊查看「頁碼與標題對照表」", expanded=False):
        st.dataframe(preview_data, use_container_width=True, height=250, hide_index=True)

    # --- 拆分任務區塊 ---
    with st.container(border=True):
        c_head1, c_head2 = st.columns([3, 1])
        c_head1.subheader("📝 步驟二：設定拆分任務")
        if c_head2.button("➕ 新增任務", type="primary", use_container_width=True):
            add_split_job(total_slides)

        if not st.session_state.split_jobs:
            st.info("☝️ 尚未建立任務，請點擊上方按鈕新增。")

        for i, job in enumerate(st.session_state.split_jobs):
            with st.container(border=True):
                st.markdown(f"**📄 任務 {i+1}**")
                
                c1, c2, c3 = st.columns([3, 1.5, 1.5])
                job["filename"] = c1.text_input("檔名", value=job["filename"], key=f"f_{job['id']}", placeholder="例如: 清潔案例A")
                job["start"] = c2.number_input("起始頁", 1, total_slides, job["start"], key=f"s_{job['id']}")
                job["end"] = c3.number_input("結束頁", 1, total_slides, job["end"], key=f"e_{job['id']}")
                
                m1, m2, m3, m4 = st.columns(4)
                job["category"] = m1.selectbox("類型", ["清潔", "配送", "購物", "AURO"], key=f"cat_{job['id']}")
                job["subcategory"] = m2.text_input("子分類", value=job["subcategory"], key=f"sub_{job['id']}")
                job["client"] = m3.text_input("客戶", value=job["client"], key=f"cli_{job['id']}")
                job["keywords"] = m4.text_input("關鍵字", value=job["keywords"], key=f"key_{job['id']}")
                
                if st.button("🗑️ 刪除此任務", key=f"d_{job['id']}", type="secondary"):
                    remove_split_job(i)
                    st.rerun()

        if st.session_state.current_file_name:
            save_history(st.session_state.current_file_name, st.session_state.split_jobs)

    # --- 執行區塊 ---
    with st.container(border=True):
        st.subheader("🚀 開始執行")
        auto_clean = st.checkbox("任務完成後自動清除暫存檔", value=True)

        if st.button("執行自動化排程", type="primary", use_container_width=True):
            if not st.session_state.split_jobs:
                st.error("請至少設定一個拆分任務！")
            else:
                validation_errors = validate_jobs(st.session_state.split_jobs, total_slides)
                if validation_errors:
                    for err in validation_errors:
                        st.error(err)
                    st.error("⛔️ 請修正錯誤後繼續。")
                else:
                    if 'bot' not in st.session_state or not st.session_state.bot:
                        st.error("❌ 機器人未初始化 (憑證錯誤)，請檢查 Secrets。")
                        st.stop()
                    
                    execute_automation_logic(
                        st.session_state.bot,
                        os.path.join(WORK_DIR, "source.pptx"),
                        os.path.splitext(st.session_state.current_file_name)[0],
                        st.session_state.split_jobs,
                        auto_clean
                    )
