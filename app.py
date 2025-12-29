import streamlit as st
import os
import uuid
import json
import shutil
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# ==========================================
#              設定頁面
# ==========================================
st.set_page_config(page_title="Aurotek數位資料庫 簡報案例自動化發布平台", layout="wide")
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# ==========================================
#              Helper Functions
# ==========================================
def cleanup_workspace():
    """強制刪除工作目錄並重建"""
    if os.path.exists(WORK_DIR):
        try:
            shutil.rmtree(WORK_DIR)
        except Exception as e:
            print(f"Cleanup warning: {e}")
    os.makedirs(WORK_DIR)

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
        task_label = f"任務 (檔名: {job['filename'] or '未命名'})"
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

# ==========================================
#              Main UI
# ==========================================
st.title("🤖 Aurotek數位資料庫 簡報案例自動化發布平台")
st.info("功能：自動拆分 PPT -> 影片雲端化 -> 內嵌優化 -> 寫入和椿數位資料庫 Google Sheets")

if 'split_jobs' not in st.session_state:
    st.session_state.split_jobs = []
if 'bot' not in st.session_state:
    try:
        if os.path.exists('credentials.json'):
            st.session_state.bot = PPTAutomationBot()
        else:
            st.error("找不到 credentials.json")
    except Exception as e:
        st.warning(f"驗證初始化中... {e}")

if 'current_file_name' not in st.session_state:
    st.session_state.current_file_name = None
if 'ppt_meta' not in st.session_state:
    st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}

# 1. 檔案上傳
uploaded_file = st.file_uploader("📂 步驟一：上傳原始 PPTX", type=['pptx'])

if uploaded_file:
    file_prefix = os.path.splitext(uploaded_file.name)[0]
    source_path = os.path.join(WORK_DIR, "source.pptx")
    
    if st.session_state.current_file_name != uploaded_file.name:
        cleanup_workspace()
        st.toast("已清除舊的暫存檔案，釋放硬碟空間。", icon="🧹")

        saved_jobs = load_history(uploaded_file.name)
        if saved_jobs:
            st.session_state.split_jobs = saved_jobs
            st.toast(f"已自動還原 {len(saved_jobs)} 筆設定！", icon="↩️")
        else:
            st.session_state.split_jobs = []

        progress_text = "正在處理大型檔案 (寫入硬碟與解析)..."
        my_bar = st.progress(0, text=progress_text)
        
        try:
            with open(source_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            my_bar.progress(50, text="寫入完成，正在解析 PPT 結構...")
            
            prs = Presentation(source_path)
            total_slides = len(prs.slides)
            
            preview_data = []
            for i, slide in enumerate(prs.slides):
                txt = "無標題"
                if slide.shapes.title and slide.shapes.title.text:
                    txt = slide.shapes.title.text
                else:
                    for s in slide.shapes:
                        if hasattr(s, "text") and s.text.strip():
                            txt = s.text.strip()[:30] + "..."
                            break
                preview_data.append({"頁碼": i+1, "內容摘要": txt})
            
            st.session_state.ppt_meta["total_slides"] = total_slides
            st.session_state.ppt_meta["preview_data"] = preview_data
            st.session_state.current_file_name = uploaded_file.name
            
            my_bar.progress(100, text="解析完成！")
            my_bar.empty()
            st.success(f"檔案讀取成功！共 {total_slides} 頁。")
            
        except Exception as e:
            st.error(f"檔案處理失敗: {e}")
            st.stop()

    total_slides = st.session_state.ppt_meta["total_slides"]
    preview_data = st.session_state.ppt_meta["preview_data"]

    # 2. 預覽
    with st.expander("👁️ 點擊查看頁碼與標題對照", expanded=True):
        st.dataframe(preview_data, use_container_width=True, height=300)

    # 3. 拆分設定
    st.divider()
    st.subheader("📝 步驟二：設定拆分任務")
    
    if st.button("➕ 新增拆分項目 (將插入至最上方)"):
        add_split_job(total_slides)

    for i, job in enumerate(st.session_state.split_jobs):
        with st.container():
            st.markdown(f"#### 🔽 任務編輯區塊") 
            c1, c2, c3, c4 = st.columns([2, 1, 1, 0.5])
            job["filename"] = c1.text_input("檔名", value=job["filename"], key=f"f_{job['id']}", placeholder="例如: MT1_Demo")
            job["start"] = c2.number_input("開始", 1, total_slides, job["start"], key=f"s_{job['id']}")
            job["end"] = c3.number_input("結束", 1, total_slides, job["end"], key=f"e_{job['id']}")
            
            if c4.button("🗑️", key=f"d_{job['id']}"):
                remove_split_job(i)
                st.rerun()
            
            m1, m2, m3, m4 = st.columns(4)
            job["category"] = m1.selectbox("Category", ["清潔", "配送", "購物", "AURO"], key=f"cat_{job['id']}")
            job["subcategory"] = m2.text_input("SubCategory", value=job["subcategory"], key=f"sub_{job['id']}")
            job["client"] = m3.text_input("Client", value=job["client"], key=f"cli_{job['id']}")
            job["keywords"] = m4.text_input("Keywords", value=job["keywords"], key=f"key_{job['id']}")
            st.markdown("---")

    if st.session_state.current_file_name:
        save_history(st.session_state.current_file_name, st.session_state.split_jobs)

    # 4. 執行選項
    st.markdown("##### ⚙️ 執行選項")
    # debug_mode = st.checkbox("🛠️ 僅產生本地拆分檔供檢查 (不上傳雲端)", value=False) # [移除] 正式版移除此選項
    auto_clean = st.checkbox("✅ 任務完成後，自動刪除所有中間暫存檔 (釋放空間)", value=True)

    # 5. 執行按鈕
    if st.button("🚀 開始自動化排程", type="primary"):
        if not st.session_state.split_jobs:
            st.error("請至少設定一個拆分任務！")
        else:
            validation_errors = validate_jobs(st.session_state.split_jobs, total_slides)
            
            if validation_errors:
                for err in validation_errors:
                    st.error(err)
                st.error("⛔️ 請修正上述錯誤後再重新開始。")
            else:
                bot = st.session_state.bot
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                step1_info = st.empty()
                step1_bar = st.empty()
                step4_info = st.empty()
                step4_bar = st.empty()

                sorted_jobs = sorted(st.session_state.split_jobs, key=lambda x: x['start'])
                
                # --- 回調函式定義區 ---
                def step1_log_handler(msg):
                    status_text.text(f"Step 1/5: {msg}")
                    st.toast(msg, icon="ℹ️")

                def update_step1_progress(filename, current, total):
                    pct = current / total
                    mb_current = current / (1024 * 1024)
                    mb_total = total / (1024 * 1024)
                    step1_info.markdown(f"**正在上傳影片: `{filename}`** ({mb_current:.2f} MB / {mb_total:.2f} MB)")
                    step1_bar.progress(pct)

                def step4_log_handler(msg):
                    status_text.text(f"Step 4/5: {msg}")

                def update_step4_progress(filename, current, total):
                    pct = current / total
                    mb_current = current / (1024 * 1024)
                    mb_total = total / (1024 * 1024)
                    step4_info.markdown(f"**正在上傳簡報: `{filename}`** ({mb_current:.2f} MB / {mb_total:.2f} MB)")
                    step4_bar.progress(pct)

                def step5_log_handler(msg):
                    status_text.text(f"Step 5/5: {msg}")
                
                def step6_log_handler(msg):
                    status_text.text(f"Final: {msg}")

                try:
                    # === Step 1 ===
                    status_text.text(f"Step 1/5: 正在提取並上傳影片...")
                    video_map = bot.extract_and_upload_videos(
                        source_path, 
                        os.path.join(WORK_DIR, "media"), 
                        file_prefix=file_prefix,
                        progress_callback=update_step1_progress,
                        log_callback=step1_log_handler
                    )
                    step1_info.empty()
                    step1_bar.empty()
                    progress_bar.progress(20)
                    
                    # === Step 2 ===
                    status_text.text("Step 2/5: 正在置換 PPT 影片...")
                    mod_path = os.path.join(WORK_DIR, "modified.pptx")
                    bot.replace_videos_with_images(source_path, mod_path, video_map)
                    progress_bar.progress(40)
                    
                    # === Step 3 ===
                    status_text.text("Step 3/5: 正在進行檔案瘦身...")
                    slim_path = os.path.join(WORK_DIR, "slim.pptx")
                    bot.shrink_pptx(mod_path, slim_path)
                    progress_bar.progress(50)
                    
                    # === Step 4 ===
                    # 正式模式，強制 debug_mode=False
                    status_text.text("Step 4/5: 正在拆分並轉換為 Google Slides...")

                    for job in sorted_jobs:
                        if not job['filename'].endswith('.pptx'):
                            job['filename'] += '.pptx'
                            
                    results = bot.split_and_upload(
                        slim_path, 
                        sorted_jobs,
                        progress_callback=update_step4_progress,
                        log_callback=step4_log_handler,
                        debug_mode=False  # <--- 強制關閉 Debug
                    )
                    
                    # 檢查錯誤
                    oversized_errors = [r for r in results if r.get('error_too_large')]
                    if oversized_errors:
                        st.error("⛔️ 偵測到檔案過大錯誤，流程已終止！")
                        for err_job in oversized_errors:
                            st.error(f"❌ 任務「{err_job['filename']}」壓縮後仍有 {err_job['size_mb']:.2f} MB，超過 Google 限制 (100MB)。")
                        st.warning("💡 請回到上方拆分設定，將上述任務拆分成更小的頁數範圍 (例如 10 頁拆成 5+5 頁)，然後重新執行。")
                        st.stop()

                    step4_info.empty()
                    step4_bar.empty()
                    progress_bar.progress(70)
                    
                    # === Step 5 ===
                    status_text.text("Step 5/5: 內嵌優化...")
                    final_results = bot.embed_videos_in_slides(
                        results,
                        log_callback=step5_log_handler
                    )
                    progress_bar.progress(85)
                    
                    # === Final ===
                    status_text.text("Final: 寫入資料庫...")
                    bot.log_to_sheets(
                        final_results,
                        log_callback=step6_log_handler
                    )
                    progress_bar.progress(100)
                    
                    status_text.success("🎉 任務完成！")
                    st.balloons()
                    
                    if auto_clean:
                        cleanup_workspace()
                        st.toast("已依您的設定清除暫存檔！", icon="🗑️")
                    
                    st.subheader("產出結果 (依頁碼順序)：")
                    for res in final_results:
                        if 'final_link' in res:
                            st.markdown(f"- **{res['filename']}**: [開啟簡報]({res['final_link']})")
                    
                except Exception as e:
                    st.error(f"執行錯誤: {e}")
                    import traceback
                    st.code(traceback.format_exc())