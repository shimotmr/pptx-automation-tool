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
#              Header 專用 function（唯一）
# ==========================================
def render_header(logo_url: str, subtitle: str, desktop_logo_px: int = 300, mobile_logo_px: int = 260):
    st.markdown(f"""
    <style>
      .auro-header {{
        display:flex;
        flex-direction:column;
        align-items:center;
        justify-content:center;
        width:100%;
        margin: 6px 0 10px 0; /* ✅ 減少上下留白 */
      }}
      .auro-header img {{
        width:{desktop_logo_px}px !important;
        height:auto !important;
        max-width:none !important; /* ✅ 避免被 global img max-width 影響 */
        display:block;
      }}
      .auro-subtitle {{
        margin-top: 4px;
        color: #6B7280;
        font-size: 1.02rem;
        font-weight: 600;
        letter-spacing: 2px;
        text-align:center;
      }}
      @media (max-width: 768px) {{
        .auro-header img {{
          width:{mobile_logo_px}px !important;
        }}
        .auro-subtitle {{
          font-size: 0.98rem;
          letter-spacing: 1px;
        }}
      }}
    </style>

    <div class="auro-header">
      <img src="{logo_url}" alt="AUROTEK LOGO" />
      <div class="auro-subtitle">{subtitle}</div>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
#              CSS 深度優化
# ==========================================
st.markdown("""
    <style>
    /* 1. 隱藏 Streamlit 預設 Header 與 Toolbar */
    header[data-testid="stHeader"] { display: none; }
    .stApp > header { display: none; }

    /* 2. 調整頂部間距（✅ 比原本更緊湊） */
    .block-container {
        padding-top: 0.9rem !important;
        padding-bottom: 1.2rem !important;
    }

    /* 3. FileUploader：瘦身 + 企業版（✅ 不用 button::after 疊字，避免直排/重複/框線錯位） */
    [data-testid="stFileUploaderDropzoneInstructions"] > div:first-child,
    [data-testid="stFileUploaderDropzoneInstructions"] > div:nth-child(2) {
        visibility: hidden; height: 0;
    }
    [data-testid="stFileUploaderDropzoneInstructions"]::before {
        content: "拖放或點擊上傳";
        visibility: visible;
        display: block;
        font-size: 0.95rem;
        font-weight: 700;
        margin-bottom: 2px;
        color: #111827;
    }
    [data-testid="stFileUploaderDropzoneInstructions"]::after {
        content: "PPTX · 單檔 5GB";
        visibility: visible;
        display: block;
        font-size: 0.75rem;
        color: #6B7280;
    }

    section[data-testid="stFileUploaderDropzone"] {
        padding: 0.7rem 1rem !important;
        border-radius: 14px !important;
        background: #F3F4F6 !important;
    }

    /* ✅ 隱藏「第二顆」瀏覽檔案按鈕（上傳後會出現的那顆） */
    div[data-testid="stFileUploader"] section:not([data-testid="stFileUploaderDropzone"]) button {
        display: none !important;
    }

    /* 4. 通用樣式 */
    h3 { font-size: 1.5rem !important; font-weight: 700 !important; }
    h4 { font-size: 1.2rem !important; font-weight: 700 !important; color: #374151; }
    .stProgress > div > div > div > div { color: white; font-weight: 600; }

    /* 5. info/callout 文字尺寸（✅ 保持清爽） */
    [data-testid="stAlert"] p {
        font-size: 0.9rem !important;
        line-height: 1.5 !important;
    }

    /* 6. Container 邊框一致化 */
    div[data-testid="stVerticalBlockBorderWrapper"] > div {
        border-radius: 16px !important;
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
            errors.append(
                f"⚠️ 發現頁數重疊："
                f"{current_job['filename']}({current_job['start']}-{current_job['end']}) 與 "
                f"{next_job['filename']}({next_job['start']}-{next_job['end']})"
            )
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

def scroll_to_bottom():
    components.html(
        "<script>window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'});</script>",
        height=0
    )

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
        detail_bar_placeholder.progress(pct, text=f"影片上傳：{filename}（{int(pct*100)}%）")
        scroll_to_bottom()

    def update_step2(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"投影片置換：{current}/{total}（{int(pct*100)}%）")
        scroll_to_bottom()

    def update_step3(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"檔案瘦身：{current}/{total}（{int(pct*100)}%）")
        scroll_to_bottom()

    def update_step4(filename, current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"拆分上傳：{filename}（{int(pct*100)}%）")
        scroll_to_bottom()

    def update_step5(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"內嵌優化：{current}/{total}（{int(pct*100)}%）")
        scroll_to_bottom()

    def general_log(msg):
        print(f"[Log] {msg}")

    try:
        status_area.info("步驟 1：提取簡報內影片並上傳雲端")
        main_progress.progress(5, text="Step 1：影片雲端化")
        video_map = bot.extract_and_upload_videos(
            source_path,
            os.path.join(WORK_DIR, "media"),
            file_prefix=file_prefix,
            progress_callback=update_step1,
            log_callback=general_log
        )
        detail_bar_placeholder.empty()

        status_area.info("步驟 2：將影片替換為雲端連結圖片")
        main_progress.progress(25, text="Step 2：連結置換")
        mod_path = os.path.join(WORK_DIR, "modified.pptx")
        bot.replace_videos_with_images(
            source_path,
            mod_path,
            video_map,
            progress_callback=update_step2
        )
        detail_bar_placeholder.empty()

        status_area.info("步驟 3：進行檔案壓縮與瘦身")
        main_progress.progress(45, text="Step 3：檔案瘦身")
        slim_path = os.path.join(WORK_DIR, "slim.pptx")
        bot.shrink_pptx(
            mod_path,
            slim_path,
            progress_callback=update_step3
        )
        detail_bar_placeholder.empty()

        status_area.info("步驟 4：依任務拆分並上傳 Google Slides")
        main_progress.progress(65, text="Step 4：拆分發布")
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
            st.error("流程終止：偵測到拆分後檔案超過 Google 100MB 限制。")
            for err_job in oversized_errors:
                st.error(f"任務「{err_job['filename']}」仍有 {err_job['size_mb']:.2f} MB")
            return

        status_area.info("步驟 5：優化線上簡報影片播放器")
        main_progress.progress(85, text="Step 5：內嵌優化")
        final_results = bot.embed_videos_in_slides(
            results,
            progress_callback=update_step5,
            log_callback=general_log
        )
        detail_bar_placeholder.empty()

        status_area.info("最後：寫入 Google Sheets 資料庫")
        main_progress.progress(95, text="Final：寫入資料庫")
        bot.log_to_sheets(
            final_results,
            log_callback=general_log
        )

        main_progress.progress(100, text="完成")
        st.info("流程已完成")

        if auto_clean:
            cleanup_workspace()
            st.toast("已自動清除暫存檔案。", icon="🧹")

        st.divider()
        st.subheader("產出結果")

        # ✅ 結果用更乾淨的呈現
        for res in final_results:
            if "final_link" in res:
                display_name = f"[{file_prefix}]_{res['filename']}"
                st.markdown(f"• **{display_name}** 　[開啟 Google Slides]({res['final_link']})")

        # ✅ 回到第一步
        if st.button("返回並處理新檔", use_container_width=True):
            st.session_state.current_file_name = None
            st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}
            st.session_state.split_jobs = []
            cleanup_workspace()
            st.rerun()

    except Exception as e:
        st.error(f"執行過程中發生錯誤: {e}")
        with st.expander("查看詳細錯誤資訊"):
            st.code(traceback.format_exc())

# ==========================================
#              Main UI
# ==========================================

# Header（✅ 唯一 LOGO render：桌機 300px，手機 260px）
render_header(LOGO_URL, "簡報案例自動化發布平台", desktop_logo_px=300, mobile_logo_px=260)

# 功能說明（保持你的藍底）
st.info("功能說明： 上傳PPT → 線上拆分 → 影片雲端化 → 內嵌優化 → 簡報雲端化 → 寫入和椿資料庫")

# 初始化
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

# Step 1：檔案來源
with st.container(border=True):
    st.subheader("步驟一：選擇檔案來源")

    input_method = st.radio("上傳方式", ["本地檔案", "線上檔案"], horizontal=True)

    uploaded_file = None
    source_path = os.path.join(WORK_DIR, "source.pptx")
    file_name_for_logic = None

    if input_method == "本地檔案":
        uploaded_file = st.file_uploader("PPTX", type=['pptx'], label_visibility="collapsed")
        if uploaded_file:
            file_name_for_logic = uploaded_file.name
            if not os.path.exists(WORK_DIR):
                os.makedirs(WORK_DIR, exist_ok=True)
            with open(source_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
    else:
        url_input = st.text_input("PPTX 直接下載網址", placeholder="https://example.com/file.pptx")
        if url_input and st.button("下載並載入", use_container_width=True):
            with st.spinner("正在下載檔案..."):
                if not os.path.exists(WORK_DIR):
                    os.makedirs(WORK_DIR, exist_ok=True)
                success, error = download_file_from_url(url_input, source_path)
                if success:
                    fake_name = url_input.split("/")[-1].split("?")[0]
                    if not fake_name.lower().endswith(".pptx"):
                        fake_name += ".pptx"
                    file_name_for_logic = fake_name
                    st.success("下載成功！")
                else:
                    st.error(f"下載失敗: {error}")

    if file_name_for_logic and os.path.exists(source_path):
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
                    preview_data.append({"頁碼": i + 1, "內容摘要": txt})

                st.session_state.ppt_meta["total_slides"] = total_slides
                st.session_state.ppt_meta["preview_data"] = preview_data
                st.session_state.current_file_name = file_name_for_logic

                progress_placeholder.progress(100, text="完成！")
                st.success(f"已讀取：{file_name_for_logic}（共 {total_slides} 頁）")

            except Exception as e:
                st.error(f"檔案處理失敗: {e}")
                st.session_state.current_file_name = None
                st.stop()

# Step 2：拆分任務（✅ 你要的欄位全部保留）
if st.session_state.current_file_name:
    total_slides = st.session_state.ppt_meta["total_slides"]
    preview_data = st.session_state.ppt_meta["preview_data"]

    with st.expander("頁碼與標題對照表", expanded=False):
        st.dataframe(preview_data, use_container_width=True, height=260, hide_index=True)

    with st.container(border=True):
        c_head1, c_head2 = st.columns([3, 1])
        c_head1.subheader("步驟二：設定拆分任務")
        if c_head2.button("新增任務", type="primary", use_container_width=True):
            add_split_job(total_slides)

        if not st.session_state.split_jobs:
            st.info("尚未建立任務，請點擊新增任務。")

        for i, job in enumerate(st.session_state.split_jobs):
            with st.container(border=True):
                st.markdown(f"**任務 {i+1}**")

                c1, c2, c3 = st.columns([3, 1.5, 1.5])
                job["filename"] = c1.text_input("檔名", value=job["filename"], key=f"f_{job['id']}", placeholder="例如：清潔案例A")
                job["start"] = c2.number_input("起始頁", 1, total_slides, job["start"], key=f"s_{job['id']}")
                job["end"] = c3.number_input("結束頁", 1, total_slides, job["end"], key=f"e_{job['id']}")

                m1, m2, m3, m4 = st.columns(4)
                job["category"] = m1.selectbox("類型", ["清潔", "配送", "購物", "AURO"], index=["清潔", "配送", "購物", "AURO"].index(job.get("category", "清潔")), key=f"cat_{job['id']}")
                job["subcategory"] = m2.text_input("子分類", value=job.get("subcategory", ""), key=f"sub_{job['id']}")
                job["client"] = m3.text_input("客戶", value=job.get("client", ""), key=f"cli_{job['id']}")
                job["keywords"] = m4.text_input("關鍵字", value=job.get("keywords", ""), key=f"key_{job['id']}")

                if st.button("刪除此任務", key=f"d_{job['id']}", type="secondary"):
                    remove_split_job(i)
                    st.rerun()

        save_history(st.session_state.current_file_name, st.session_state.split_jobs)

    # Step 3：執行
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
                else:
                    if 'bot' not in st.session_state or not st.session_state.bot:
                        st.error("機器人未初始化（憑證錯誤），請檢查 Secrets。")
                        st.stop()

                    execute_automation_logic(
                        st.session_state.bot,
                        os.path.join(WORK_DIR, "source.pptx"),
                        os.path.splitext(st.session_state.current_file_name)[0],
                        st.session_state.split_jobs,
                        auto_clean
                    )
