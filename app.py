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
    page_icon="🤖",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# ==========================================
# CSS（保持您的企業版風格，並微調 Logo 容器）
# ==========================================
st.markdown("""
<style>
header[data-testid="stHeader"] { display:none; }
.block-container { padding-top:1rem; padding-bottom: 2rem; }

/* Logo 容器：Flexbox 置中 */
.auro-header{
  display:flex;
  flex-direction:column;
  align-items:center;
  justify-content: center;
  margin-bottom: 20px;
}
.auro-header img{ 
    width: 300px !important; 
    max-width: 90vw !important; 
    height: auto; 
}
.auro-sub{ 
    color:#6B7280; 
    font-weight:600; 
    letter-spacing:2px; 
    margin-top: 5px;
    font-size: 1rem;
}

/* Callout 風格 */
.callout{
  border:1px solid #E5E7EB;
  border-left:4px solid #0B4F8A;
  background:#F9FAFB; /* 稍微淡一點的灰 */
  padding:15px;
  border-radius:8px;
  font-size: 0.95rem;
  color: #374151;
  line-height: 1.5;
}
.callout.err{
  border-left-color:#B91C1C;
  background:#FEF2F2;
  color:#991B1B;
}

/* 區塊風格 */
.section{
  border:1px solid #E5E7EB;
  border-radius:12px;
  padding:20px;
  margin-bottom:20px;
  background: white;
  box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
}

/* 進度條優化 */
.stProgress > div > div > div > div { color: white; font-weight: 500; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# Helper：工作區
# ==========================================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)

def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        try:
            shutil.rmtree(WORK_DIR)
        except:
            pass
    ensure_workspace()

# ==========================================
# Helper：網路下載 (新增)
# ==========================================
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
# Helper：歷史任務
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
# Helper：安全影片替換
# ==========================================
def safe_replace_videos(bot, source_path, output_path, video_map, progress_cb=None):
    if video_map and isinstance(video_map, dict) and len(video_map) > 0:
        bot.replace_videos_with_images(
            source_path,
            output_path,
            video_map,
            progress_callback=progress_cb
        )
        return "replaced"
    else:
        shutil.copyfile(source_path, output_path)
        return "skipped"

# ==========================================
# Header (Flexbox 300px 置中)
# ==========================================
st.markdown(f"""
<div class="auro-header">
  <img src="{LOGO_URL}">
  <div class="auro-sub">簡報案例自動化發布平台</div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="callout">
<b>功能流程：</b>上傳簡報 → 拆分任務 → 影片雲端化（如有） → 簡報優化 → Google Slides 發布 → 寫入資料庫
</div>
<div style="height:10px;"></div>
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
    try:
        st.session_state.bot = PPTAutomationBot()
    except:
        st.warning("⚠️ Bot 初始化失敗，請檢查憑證設定。")

# ==========================================
# Step 1：上傳檔案 (整合本地與網址)
# ==========================================
with st.container():
    st.markdown("<div class='section'>", unsafe_allow_html=True)
    st.subheader("步驟一：選擇檔案")

    ensure_workspace()
    source_path = os.path.join(WORK_DIR, "source.pptx")
    
    # 選擇來源
    input_method = st.radio("上傳方式", ["本地檔案", "線上檔案"], horizontal=True, label_visibility="collapsed")

    file_ready = False
    new_file_name = None

    if input_method == "本地檔案":
        uploaded = st.file_uploader("選擇 PPTX 檔案", type=["pptx"])
        if uploaded:
            new_file_name = uploaded.name
            with open(source_path, "wb") as f:
                f.write(uploaded.getbuffer())
            file_ready = True
    else:
        c1, c2 = st.columns([3, 1])
        url = c1.text_input("輸入 PPTX 下載網址", placeholder="https://example.com/file.pptx")
        if c2.button("下載檔案", use_container_width=True):
            if url:
                with st.spinner("下載中..."):
                    ok, err = download_file_from_url(url, source_path)
                    if ok:
                        # 簡單從網址取檔名
                        fname = url.split("/")[-1].split("?")[0]
                        if not fname.lower().endswith(".pptx"): fname += ".pptx"
                        new_file_name = fname
                        file_ready = True
                        st.success("下載成功！")
                    else:
                        st.error(f"下載失敗: {err}")

    # 檔案就緒後的處理 (解析 PPT)
    if file_ready and new_file_name:
        # 如果是新檔案，重置狀態
        if st.session_state.current_file != new_file_name:
            cleanup_workspace() # 清理舊檔
            # 確保 source_path 還在 (因為 cleanup 可能會刪除，這裡要小心)
            # 由於我們剛寫入，cleanup 應該在寫入前做，但為了邏輯簡單，我們假設 cleanup 只清 temp_workspace 下的其他東西
            # 更好的做法：
            pass 

        # 解析 PPT 資訊
        try:
            prs = Presentation(source_path)
            total = len(prs.slides)
            preview = []
            for i, s in enumerate(prs.slides):
                t = s.shapes.title.text if s.shapes.title and s.shapes.title.text else "無標題"
                # 如果沒有標題，嘗試抓第一個文字框
                if t == "無標題":
                    for shape in s.shapes:
                        if hasattr(shape, "text") and shape.text.strip():
                            t = shape.text.strip()[:20] + "..."
                            break
                preview.append({"頁碼": i + 1, "標題": t})

            st.session_state.current_file = new_file_name
            st.session_state.ppt_meta = {"total": total, "preview": preview}
            
            # 嘗試載入歷史設定
            saved_jobs = load_history(new_file_name)
            if saved_jobs:
                st.session_state.split_jobs = saved_jobs
            elif st.session_state.current_file != new_file_name:
                st.session_state.split_jobs = []

            st.success(f"✅ 已讀取 {new_file_name}（共 {total} 頁）")

        except Exception as e:
            st.error(f"檔案解析失敗: {e}")

    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# Step 2：拆分任務
# ==========================================
if st.session_state.current_file:
    with st.expander("👁️ 查看頁碼對照表", expanded=False):
        st.dataframe(
            st.session_state.ppt_meta["preview"],
            use_container_width=True,
            hide_index=True
        )

    with st.container():
        st.markdown("<div class='section'>", unsafe_allow_html=True)
        c_head1, c_head2 = st.columns([3, 1])
        c_head1.subheader("步驟二：設定拆分任務")
        if c_head2.button("➕ 新增任務", use_container_width=True):
            st.session_state.split_jobs.append({
                "id": str(uuid.uuid4())[:8],
                "filename": "",
                "start": 1,
                "end": st.session_state.ppt_meta["total"],
                "category": "清潔",
                "subcategory": "",
                "client": "",
                "keywords": ""
            })

        if not st.session_state.split_jobs:
            st.info("尚未建立任務，請點擊右上方按鈕新增。")

        for i, job in enumerate(st.session_state.split_jobs):
            with st.container():
                st.markdown(f"**📄 任務 {i+1}**")
                c1, c2, c3, c4 = st.columns([3, 1.2, 1.2, 0.5])
                job["filename"] = c1.text_input("檔名", job["filename"], key=f"f{i}", placeholder="例: Case_A")
                job["start"] = c2.number_input("起始頁", 1, st.session_state.ppt_meta["total"], job["start"], key=f"s{i}")
                job["end"] = c3.number_input("結束頁", 1, st.session_state.ppt_meta["total"], job["end"], key=f"e{i}")
                
                if c4.button("🗑️", key=f"del{i}"):
                    st.session_state.split_jobs.pop(i)
                    st.rerun()

                m1, m2, m3, m4 = st.columns(4)
                job["category"] = m1.selectbox("類型", ["清潔", "配送", "購物", "AURO"], index=0, key=f"c{i}")
                job["subcategory"] = m2.text_input("子分類", job["subcategory"], key=f"sc{i}")
                job["client"] = m3.text_input("客戶", job["client"], key=f"cl{i}")
                job["keywords"] = m4.text_input("關鍵字", job["keywords"], key=f"k{i}")
                st.markdown("---")

        save_history(st.session_state.current_file, st.session_state.split_jobs)
        st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# Step 3：執行 (整合進度條)
# ==========================================
if st.session_state.current_file:
    with st.container():
        st.markdown("<div class='section'>", unsafe_allow_html=True)
        st.subheader("步驟三：開始執行")
        
        auto_clean = st.checkbox("完成後自動清理暫存檔", value=True)

        if st.button("🚀 執行自動化排程", type="primary", use_container_width=True):
            if not st.session_state.split_jobs:
                st.error("請至少設定一個任務！")
            else:
                try:
                    bot = st.session_state.bot
                    
                    # 準備 UI 元件
                    main_bar = st.progress(0, text="準備中...")
                    status_text = st.empty()
                    detail_bar_placeholder = st.empty()

                    # 定義回調函數
                    def update_step1(fname, curr, tot):
                        p = curr / tot if tot else 0
                        detail_bar_placeholder.progress(p, text=f"正在上傳影片: {fname}")

                    def update_step2(curr, tot):
                        p = curr / tot if tot else 0
                        detail_bar_placeholder.progress(p, text=f"替換連結中: {curr}/{tot}")

                    def update_step3(curr, tot):
                        p = curr / tot if tot else 0
                        detail_bar_placeholder.progress(p, text=f"圖片壓縮中: {curr}/{tot}")

                    def update_step4(fname, curr, tot):
                        p = curr / tot if tot else 0
                        detail_bar_placeholder.progress(p, text=f"上傳簡報: {fname}")

                    def update_step5(curr, tot):
                        p = curr / tot if tot else 0
                        detail_bar_placeholder.progress(p, text=f"優化內嵌: {curr}/{tot}")

                    def log_handler(msg):
                        print(f"[Log] {msg}")

                    # Step 1: 影片
                    status_text.info("1️⃣ 處理影片...")
                    main_bar.progress(10, text="Step 1: 影片雲端化")
                    
                    video_map_path = os.path.join(WORK_DIR, "video_map.json")
                    if os.path.exists(video_map_path):
                        with open(video_map_path, "r") as f: video_map = json.load(f)
                    else:
                        video_map = bot.extract_and_upload_videos(
                            source_path,
                            os.path.join(WORK_DIR, "media"),
                            file_prefix=os.path.splitext(st.session_state.current_file)[0],
                            progress_callback=update_step1,
                            log_callback=log_handler
                        )
                        with open(video_map_path, "w") as f: json.dump(video_map, f)
                    
                    detail_bar_placeholder.empty()

                    # Step 2: 替換
                    status_text.info("2️⃣ 替換連結...")
                    main_bar.progress(30, text="Step 2: 連結替換")
                    modified = os.path.join(WORK_DIR, "modified.pptx")
                    res = safe_replace_videos(bot, source_path, modified, video_map, progress_cb=update_step2)
                    if res == "skipped": st.caption("無影片，已略過此步驟。")
                    detail_bar_placeholder.empty()

                    # Step 3: 瘦身
                    status_text.info("3️⃣ 檔案瘦身...")
                    main_bar.progress(50, text="Step 3: 圖片壓縮")
                    slim = os.path.join(WORK_DIR, "slim.pptx")
                    bot.shrink_pptx(modified, slim, progress_callback=update_step3)
                    detail_bar_placeholder.empty()

                    # Step 4: 拆分上傳
                    status_text.info("4️⃣ 拆分與發布...")
                    main_bar.progress(70, text="Step 4: 拆分上傳")
                    results = bot.split_and_upload(
                        slim,
                        st.session_state.split_jobs,
                        file_prefix=os.path.splitext(st.session_state.current_file)[0],
                        progress_callback=update_step4,
                        log_callback=log_handler
                    )
                    detail_bar_placeholder.empty()

                    # Step 5: 內嵌
                    status_text.info("5️⃣ 優化線上播放器...")
                    main_bar.progress(90, text="Step 5: 內嵌優化")
                    final = bot.embed_videos_in_slides(results, progress_callback=update_step5, log_callback=log_handler)
                    detail_bar_placeholder.empty()

                    # Final: 寫入
                    status_text.info("📝 寫入資料庫...")
                    bot.log_to_sheets(final, log_callback=log_handler)

                    main_bar.progress(100, text="🎉 完成！")
                    status_text.success("所有任務執行完畢！")
                    st.balloons()

                    if auto_clean:
                        cleanup_workspace()
                        st.toast("暫存檔已清理", icon="🧹")

                    # 顯示結果
                    st.divider()
                    st.subheader("✅ 產出連結")
                    cnt = 0
                    for r in final:
                        if "final_link" in r:
                            cnt += 1
                            dname = f"[{os.path.splitext(st.session_state.current_file)[0]}]_{r['filename']}"
                            st.markdown(f"👉 **{dname}**: [開啟 Google Slides]({r['final_link']})")
                    
                    if cnt == 0:
                        st.warning("沒有產生任何連結，請檢查日誌。")

                except Exception as e:
                    st.markdown(f"<div class='callout err'>發生錯誤：{e}</div>", unsafe_allow_html=True)
                    st.code(traceback.format_exc())

        st.markdown("</div>", unsafe_allow_html=True)