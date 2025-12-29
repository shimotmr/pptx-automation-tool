# Version: v0.6
# Update Log:
# 1. Added "Start New Project" (Reset) button after completion.
# 2. Result list is now a structured table with "Open" and "Copy" functions.
# 3. Removed all "fun" emojis (balloons, rockets) for a professional corporate look.
# 4. Standardized font sizes and colors across all sections.

import streamlit as st
import streamlit.components.v1 as components
import os
import uuid
import json
import shutil
import traceback
import requests
import pandas as pd # 引入 pandas 以防萬一需要，雖然此版主要用 columns 模擬表格
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# ==========================================
#              設定頁面與樣式
# ==========================================
st.set_page_config(
    page_title="Aurotek數位資料庫 簡報案例自動化發布平台",
    page_icon="📄", # 改為文件圖示，較正式
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# ==========================================
#              CSS 深度優化 (企業風格)
# ==========================================
st.markdown("""
<style>
/* 1. 隱藏 Streamlit 預設 Header 與 Toolbar */
header[data-testid="stHeader"] { display: none; }
.stApp > header { display: none; }

/* 2. 調整頂部間距 */
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 3rem !important;
}

/* 3. 上傳元件中文化 */
[data-testid="stFileUploaderDropzoneInstructions"] > div:first-child { display: none !important; }
[data-testid="stFileUploaderDropzoneInstructions"] > div:nth-child(2) { display: none !important; }

[data-testid="stFileUploaderDropzoneInstructions"]::before {
    content: "請將檔案拖放至此";
    display: block;
    font-size: 0.95rem;
    font-weight: 700;
    margin: 0;
    line-height: 1.2;
    color: #31333F;
}
[data-testid="stFileUploaderDropzoneInstructions"]::after {
    content: "單一檔案限制 5GB • PPTX";
    display: block;
    font-size: 0.75rem;
    color: #8a8a8a;
    margin-top: 4px;
    line-height: 1.2;
}

/* 瀏覽檔案按鈕修正 */
[data-testid="stFileUploader"] button { 
    font-size: 0 !important;
    line-height: 0 !important;
    color: transparent !important;
    position: relative;
    width: auto !important;
    min-width: 100px !important; 
    height: auto !important;
    min-height: 38px !important;
    padding: 0 15px !important;
    overflow: visible !important;
    border: 1px solid #e0e0e0;
}

[data-testid="stFileUploader"] button::after {
    content: "瀏覽檔案";
    font-size: 0.9rem !important;
    line-height: 1.5 !important;
    color: #31333F !important;
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    font-weight: 500;
    white-space: nowrap;
    display: block;
}

/* 4. 統一字體與標題樣式 (去除誇張大小) */
h3 { font-size: 1.2rem !important; font-weight: 600 !important; color: #31333F; margin-bottom: 0.5rem;}
h4 { font-size: 1.1rem !important; font-weight: 600 !important; color: #555; }
.stProgress > div > div > div > div { color: white; font-weight: 500; }

/* 5. 統一提示詞顏色 (全部藍色風格) */
div[data-testid="stAlert"][data-style="success"],
div[data-testid="stAlert"][data-style="info"] {
    background-color: #F0F2F6 !important; /* 統一淺灰藍底 */
    color: #31333F !important; /* 深灰文字 */
    border: 1px solid #d0d7de !important;
}
/* 圖示顏色統一 */
div[data-testid="stAlert"] svg {
    color: #004280 !important; 
}

/* 縮小提示框文字 */
[data-testid="stAlert"] p {
    font-size: 0.9rem !important;
    line-height: 1.4 !important;
}

/* 結果表格樣式微調 */
div[data-testid="column"] {
    align-content: center;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
#              Helper Functions
# ==========================================
def cleanup_workspace():
    """清理工作目錄"""
    if os.path.exists(WORK_DIR):
        try:
            shutil.rmtree(WORK_DIR)
        except Exception as e:
            print(f"Cleanup warning: {e}")
    os.makedirs(WORK_DIR, exist_ok=True)

def reset_session():
    """重置所有狀態，準備新任務"""
    cleanup_workspace()
    st.session_state.split_jobs = []
    st.session_state.current_file_name = None
    st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}
    st.rerun()

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
        display_num = len(jobs) - i
        task_label = f"任務 {display_num} (檔名: {job['filename'] or '未命名'})"
        
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
        detail_bar_placeholder.progress(pct, text=f"Step 1 進度: 上傳影片 `{filename}` ({int(pct*100)}%)")

    def update_step2(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"Step 2 進度: 處理投影片 {current}/{total} ({int(pct*100)}%)")

    def update_step3(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"Step 3 進度: 處理內部檔案 {current}/{total} ({int(pct*100)}%)")

    def update_step4(filename, current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"Step 4 進度: 上傳簡報 `{filename}` ({int(pct*100)}%)")

    def update_step5(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"Step 5 進度: 優化任務 {current}/{total} ({int(pct*100)}%)")

    def general_log(msg):
        print(f"[Log] {msg}")

    try:
        status_area.info("執行中：Step 1/5 - 提取影片並上傳雲端...", icon="⏳")
        main_progress.progress(5, text="Step 1: 影片雲端化")
        video_map = bot.extract_and_upload_videos(
            source_path,
            os.path.join(WORK_DIR, "media"),
            file_prefix=file_prefix,
            progress_callback=update_step1,
            log_callback=general_log
        )
        detail_bar_placeholder.empty()

        status_area.info("執行中：Step 2/5 - 替換影片連結...", icon="⏳")
        main_progress.progress(25, text="Step 2: 連結置換")
        mod_path = os.path.join(WORK_DIR, "modified.pptx")
        bot.replace_videos_with_images(
            source_path,
            mod_path,
            video_map,
            progress_callback=update_step2
        )
        detail_bar_placeholder.empty()

        status_area.info("執行中：Step 3/5 - 檔案壓縮優化...", icon="⏳")
        main_progress.progress(45, text="Step 3: 檔案瘦身")
        slim_path = os.path.join(WORK_DIR, "slim.pptx")
        bot.shrink_pptx(
            mod_path,
            slim_path,
            progress_callback=update_step3
        )
        detail_bar_placeholder.empty()

        status_area.info("執行中：Step 4/5 - 拆分並發布至 Google Slides...", icon="⏳")
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
            st.error("流程終止：偵測到檔案過大。")
            for err_job in oversized_errors:
                st.error(f"任務「{err_job['filename']}」壓縮後仍有 {err_job['size_mb']:.2f} MB，超過 Google 限制 (100MB)。")
            st.warning("建議做法：請減少該任務的頁數範圍，重新執行。")
            return

        status_area.info("執行中：Step 5/5 - 優化線上播放器...", icon="⏳")
        main_progress.progress(85, text="Step 5: 內嵌優化")
        final_results = bot.embed_videos_in_slides(
            results,
            progress_callback=update_step5,
            log_callback=general_log
        )
        detail_bar_placeholder.empty()

        status_area.info("執行中：最後步驟 - 寫入資料庫...", icon="⏳")
        main_progress.progress(95, text="Final: 寫入資料庫")
        bot.log_to_sheets(final_results, log_callback=general_log)

        main_progress.progress(100, text="任務完成")
        status_area.info("所有自動化流程執行完畢。", icon="✅")
        # [UI修正] 移除氣球 st.balloons()，保持專業

        if auto_clean:
            cleanup_workspace()
            
        st.divider()
        st.subheader("產出結果清單")
        
        # [UI修正] 結果列表表格化
        # 表頭
        cols = st.columns([4, 2, 2])
        cols[0].markdown("**檔案名稱**")
        cols[1].markdown("**線上預覽**")
        cols[2].markdown("**複製連結**")
        st.markdown("---")

        result_count = 0
        for res in final_results:
            if 'final_link' in res:
                result_count += 1
                display_name = f"[{file_prefix}]_{res['filename']}"
                link = res['final_link']
                
                row = st.columns([4, 2, 2])
                # 欄位 1: 檔名
                row[0].text(display_name)
                # 欄位 2: 開啟按鈕
                row[1].link_button("開啟簡報", link)
                # 欄位 3: 複製功能 (使用 st.code 實現複製按鈕)
                row[2].code(link, language=None)
        
        if result_count == 0:
            st.warning("沒有產生任何結果，請檢查是否有任務被跳過。")
        else:
            st.markdown("---")
            # [UI修正] 新增重置按鈕
            if st.button("開始新專案 (清除資料)", type="secondary", use_container_width=True):
                reset_session()

    except Exception as e:
        st.error(f"執行過程中發生錯誤: {e}")
        with st.expander("查看詳細錯誤資訊"):
            st.code(traceback.format_exc())

# ==========================================
#              Main UI (Layout)
# ==========================================

os.makedirs(WORK_DIR, exist_ok=True)

# 1) Header
st.markdown(
    f"""
    <div style="
        width: 100%;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        margin: 4px 0 20px 0;
        line-height: 1.1;
    ">
        <img src="{LOGO_URL}" alt="Aurotek Logo" style="
            width: 300px;
            max-width: 90vw;
            height: auto;
            display: block;
            margin: 0;
        " />
        <div style="
            margin-top: 10px;
            width: 100%;
            text-align: center;
            color: gray;
            font-size: 1.0rem;
            font-weight: 500;
            letter-spacing: 2px;
        ">
            簡報案例自動化發布平台
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# 2. 功能說明
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
    st.subheader("步驟一：選擇檔案來源")

    input_method = st.radio("上傳方式", ["本地檔案", "線上檔案"], horizontal=True)

    uploaded_file = None
    source_path = os.path.join(WORK_DIR, "source.pptx")
    file_name_for_logic = None

    # --- 本地檔案上傳 ---
    if input_method == "本地檔案":
        uploaded_file = st.file_uploader("請選擇 PPTX 檔案", type=['pptx'], label_visibility="collapsed")
        if uploaded_file:
            file_name_for_logic = uploaded_file.name
            
            # 新檔案 -> 清理 -> 寫入
            if st.session_state.current_file_name != file_name_for_logic:
                cleanup_workspace()
                with open(source_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
            elif not os.path.exists(source_path):
                 with open(source_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

    # --- 線上檔案下載 ---
    else:
        url_input = st.text_input("請輸入 PPTX 檔案的直接下載網址 (Direct URL)", placeholder="https://example.com/file.pptx")
        if url_input:
            if not url_input.lower().endswith(".pptx"):
                st.warning("⚠️ 網址結尾似乎不是 .pptx，請確認網址正確性。")

            fake_name = url_input.split("/")[-1].split("?")[0]
            if not fake_name.lower().endswith(".pptx"):
                fake_name += ".pptx"

            if st.button("下載並處理此網址"):
                with st.spinner("正在從網址下載檔案..."):
                    cleanup_workspace()
                    success, error = download_file_from_url(url_input, source_path)
                    if success:
                        file_name_for_logic = fake_name
                        st.info("下載成功", icon="✅")
                    else:
                        st.error(f"下載失敗: {error}")

    # 5. 檔案處理邏輯
    if file_name_for_logic and os.path.exists(source_path):
        file_prefix = os.path.splitext(file_name_for_logic)[0]

        if st.session_state.current_file_name != file_name_for_logic:
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
                st.info(f"已讀取：{file_name_for_logic} (共 {total_slides} 頁)", icon="✅")

            except Exception as e:
                st.error(f"檔案處理失敗: {e}")
                st.session_state.current_file_name = None
                st.stop()

if st.session_state.current_file_name:
    total_slides = st.session_state.ppt_meta["total_slides"]
    preview_data = st.session_state.ppt_meta["preview_data"]

    with st.expander("點擊查看「頁碼與標題對照表」", expanded=False):
        st.dataframe(preview_data, use_container_width=True, height=250, hide_index=True)

    # --- 拆分任務區塊 ---
    with st.container(border=True):
        c_head1, c_head2 = st.columns([3, 1])
        c_head1.subheader("步驟二：設定拆分任務")
        if c_head2.button("新增任務", type="primary", use_container_width=True):
            add_split_job(total_slides)

        if not st.session_state.split_jobs:
            st.info("尚未建立任務，請點擊上方按鈕新增。")

        # 計算總任務數
        total_jobs_count = len(st.session_state.split_jobs)

        for i, job in enumerate(st.session_state.split_jobs):
            display_number = total_jobs_count - i
            
            with st.container(border=True):
                st.markdown(f"**任務 {display_number}**")

                c1, c2, c3 = st.columns([3, 1.5, 1.5])
                job["filename"] = c1.text_input("檔名", value=job["filename"], key=f"f_{job['id']}", placeholder="例如: 清潔案例A")
                job["start"] = c2.number_input("起始頁", 1, total_slides, job["start"], key=f"s_{job['id']}")
                job["end"] = c3.number_input("結束頁", 1, total_slides, job["end"], key=f"e_{job['id']}")

                m1, m2, m3, m4 = st.columns(4)
                job["category"] = m1.selectbox("類型", ["清潔", "配送", "購物", "AURO"], key=f"cat_{job['id']}")
                job["subcategory"] = m2.text_input("子分類", value=job["subcategory"], key=f"sub_{job['id']}")
                job["client"] = m3.text_input("客戶", value=job["client"], key=f"cli_{job['id']}")
                job["keywords"] = m4.text_input("關鍵字", value=job["keywords"], key=f"key_{job['id']}")

                if st.button("刪除此任務", key=f"d_{job['id']}", type="secondary"):
                    remove_split_job(i)
                    st.rerun()

        if st.session_state.current_file_name:
            save_history(st.session_state.current_file_name, st.session_state.split_jobs)

    # --- 執行區塊 ---
    with st.container(border=True):
        st.subheader("步驟三：開始執行")
        auto_clean = st.checkbox("任務完成後自動清除暫存檔", value=True)

        if st.button("執行自動化排程", type="primary", use_container_width=True):
            if not st.session_state.split_jobs:
                st.error("請至少設定一個拆分任務！")
            else:
                validation_errors = validate_jobs(st.session_state.split_jobs, total_slides)
                if validation_errors:
                    for err in validation_errors:
                        st.error(err)
                    st.error("請修正錯誤後繼續。")
                else:
                    if 'bot' not in st.session_state or not st.session_state.bot:
                        st.error("機器人未初始化 (憑證錯誤)，請檢查 Secrets。")
                        st.stop()

                    execute_automation_logic(
                        st.session_state.bot,
                        os.path.join(WORK_DIR, "source.pptx"),
                        os.path.splitext(st.session_state.current_file_name)[0],
                        st.session_state.split_jobs,
                        auto_clean
                    )