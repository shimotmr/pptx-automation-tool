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
    page_title="Aurotek｜簡報案例自動化發布平台",
    page_icon="📊",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# ==========================================
#              企業版 CSS（保留功能、重做風格）
# ==========================================
st.markdown("""
<style>
/* ---- 隱藏 Streamlit 預設 Header ---- */
header[data-testid="stHeader"] { display: none; }
.stApp > header { display: none; }

/* ---- 版面留白（減少 LOGO 上下空白） ---- */
.block-container {
  padding-top: 0.9rem !important;
  padding-bottom: 1.0rem !important;
}

/* ---- 統一字級 ---- */
h3 { font-size: 1.35rem !important; font-weight: 700 !important; }
h4 { font-size: 1.05rem !important; font-weight: 650 !important; color: #1f2937; }
[data-testid="stAlert"] p { font-size: 0.90rem !important; line-height: 1.45 !important; }

/* ---- 品牌色 ---- */
:root{
  --brand-blue:#0B4F8A;
  --brand-blue-weak:#EAF3FF;
  --border:#E5E7EB;
  --text:#111827;
  --muted:#6B7280;
  --bg-soft:#F8FAFC;
}

/* ---- Header（LOGO + 副標） ---- */
.auro-header {
  display:flex;
  flex-direction:column;
  align-items:center;
  justify-content:center;
  margin: 0 0 8px 0;
}
.auro-header img{
  width:300px;
  height:auto;
}
.auro-subtitle{
  margin-top:4px;
  color: var(--muted);
  font-size: 1.00rem;
  font-weight: 600;
  letter-spacing: 2px;
  text-align:center;
}

/* ---- Callout（取代綠色 success）---- */
.callout{
  border:1px solid var(--border);
  border-radius:12px;
  padding:12px 14px;
  margin: 10px 0;
  background: #fff;
}
.callout.blue{
  border-left: 4px solid var(--brand-blue);
  background: var(--brand-blue-weak);
  color: var(--brand-blue);
  font-weight: 650;
}
.callout.gray{
  background: var(--bg-soft);
  color: var(--text);
}
.callout.warn{
  border-left: 4px solid #B45309;
  background:#FFF7ED;
  color:#92400E;
  font-weight:650;
}
.callout.err{
  border-left: 4px solid #B91C1C;
  background:#FEF2F2;
  color:#991B1B;
  font-weight:650;
}

/* ---- 卡片容器（你原本 st.container(border=True) 的企業版外觀）---- */
.section-card{
  border:1px solid var(--border);
  border-radius:16px;
  padding: 14px 14px 6px 14px;
  background:#fff;
}

/* ---- 進度條字色 ---- */
.stProgress > div > div > div > div { color: white; font-weight: 600; }

/* ==========================================
   FileUploader：修正「瀏覽檔案」重複 / 縱排 / 框線錯位
   核心做法：
   1) 只改 dropzone 內那顆按鈕（避免影響其他按鈕）
   2) 隱藏「檔案列表右側」那顆重複按鈕
   3) 用 font-size:0 取代 color:transparent，避免文字殘影/換行
========================================== */

/* 隱藏原本的兩行長說明 */
[data-testid="stFileUploaderDropzoneInstructions"] > div:first-child { display:none !important; }
[data-testid="stFileUploaderDropzoneInstructions"] > div:nth-child(2) { display:none !important; }

/* 自訂更精簡文案（不佔空間） */
[data-testid="stFileUploaderDropzoneInstructions"]::before{
  content:"拖放或點擊上傳";
  display:block;
  font-size:0.92rem;
  font-weight:750;
  line-height:1.2;
  margin:0;
}
[data-testid="stFileUploaderDropzoneInstructions"]::after{
  content:"PPTX · 單檔 5GB";
  display:block;
  font-size:0.74rem;
  color: var(--muted);
  line-height:1.15;
  margin-top:2px;
}

/* 壓縮 dropzone 高度 */
section[data-testid="stFileUploaderDropzone"]{
  padding: 0.60rem 0.90rem !important;
  border-radius:14px !important;
  background: var(--bg-soft) !important;
}

/* 只針對 dropzone 內的 button 做中文化（避免影響別的 button） */
section[data-testid="stFileUploaderDropzone"] button{
  font-size:0 !important;     /* 隱藏原文字 */
  white-space:nowrap !important;
  display:flex !important;
  align-items:center !important;
  justify-content:center !important;
  min-height:42px !important;
  line-height:1 !important;
  border-radius:12px !important;
  padding: 0 14px !important;
}
section[data-testid="stFileUploaderDropzone"] button::after{
  content:"瀏覽檔案";
  font-size:0.92rem;
  font-weight:750;
  color:#111827;
}

/* 隱藏「檔案列表右側」那顆重複的按鈕（你截圖右邊又出現一次那顆） */
div[data-testid="stFileUploader"] section:not([data-testid="stFileUploaderDropzone"]) button{
  display:none !important;
}

/* 手機更緊湊 */
@media (max-width: 768px){
  .block-container { padding-top:0.7rem !important; }
  .auro-header img { width: 280px; }
  .auro-subtitle { font-size:0.98rem; letter-spacing:1px; }
}
</style>
""", unsafe_allow_html=True)

# ==========================================
#              Helper Functions
# ==========================================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)

def cleanup_workspace():
    """完全清除工作目錄（注意：不要在寫入 source.pptx 後立刻呼叫）"""
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
        task_label = f"任務 {i+1}（{job['filename'] or '未命名'}）"
        if not job['filename'].strip():
            errors.append(f"❌ {task_label}：檔名不能為空。")
        if job['start'] > job['end']:
            errors.append(f"❌ {task_label}：起始頁不能大於結束頁。")
        if job['end'] > total_slides:
            errors.append(f"❌ {task_label}：結束頁超出簡報總頁數（{total_slides}）。")

    sorted_jobs = sorted(jobs, key=lambda x: x['start'])
    for i in range(len(sorted_jobs) - 1):
        current_job = sorted_jobs[i]
        next_job = sorted_jobs[i+1]
        if current_job['end'] >= next_job['start']:
            errors.append(
                f"⚠️ 頁數重疊：{current_job['filename']}（{current_job['start']}-{current_job['end']}）"
                f" 與 {next_job['filename']}（{next_job['start']}-{next_job['end']}）"
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

def reset_to_step1(keep_bot=True):
    """一鍵回到第一步（保留 bot 憑證，避免重登）"""
    keys = [
        "current_file_name", "ppt_meta", "split_jobs", "total_slides",
    ]
    for k in keys:
        if k in st.session_state:
            del st.session_state[k]
    if not keep_bot and "bot" in st.session_state:
        del st.session_state["bot"]
    cleanup_workspace()
    st.rerun()

# ==========================================
#              Results UI (企業版)
# ==========================================
def render_result_cards(file_prefix, final_results):
    items = []
    for res in final_results:
        link = res.get("final_link")
        if not link:
            continue
        display_name = f"[{file_prefix}]_{res['filename']}"
        items.append((display_name, link))

    if not items:
        st.markdown("<div class='callout warn'>未產生任何結果連結，請檢查是否有任務被跳過。</div>", unsafe_allow_html=True)
        return

    st.subheader("產出結果")
    # 使用 components.html：確保複製功能可靠（可執行 JS）
    cards_html = """
    <style>
      .wrap{font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Noto Sans TC","PingFang TC",Arial;}
      .card{
        border:1px solid #E5E7EB;
        border-radius:14px;
        padding:12px 14px;
        margin:10px 0;
        display:flex;
        align-items:center;
        justify-content:space-between;
        background:#fff;
      }
      .left{display:flex;flex-direction:column;gap:4px;}
      .title{font-weight:750;color:#111827;font-size:14px;}
      .meta{font-size:12px;color:#6B7280;}
      .actions{display:flex;align-items:center;gap:10px;flex-wrap:nowrap;}
      .open{
        text-decoration:none;
        background:#EAF3FF;
        color:#0B4F8A;
        padding:8px 10px;
        border-radius:10px;
        font-weight:750;
        font-size:13px;
        border:1px solid #D6E8FF;
        white-space:nowrap;
      }
      .copy{
        border:1px solid #E5E7EB;
        background:#F8FAFC;
        border-radius:10px;
        padding:8px 10px;
        cursor:pointer;
        font-weight:750;
        font-size:13px;
        white-space:nowrap;
      }
      .toast{
        position:fixed;
        right:18px;
        bottom:18px;
        background:#0B4F8A;
        color:#fff;
        padding:10px 12px;
        border-radius:12px;
        font-weight:700;
        font-size:13px;
        opacity:0;
        transform: translateY(6px);
        transition: all .18s ease;
        z-index:9999;
      }
      .toast.show{
        opacity:1;
        transform: translateY(0px);
      }
    </style>
    <div class="wrap">
    """

    for name, link in items:
        safe_name = name.replace('"', '\\"')
        safe_link = link.replace('"', '\\"')
        cards_html += f"""
        <div class="card">
          <div class="left">
            <div class="title">{safe_name}</div>
            <div class="meta">Google Slides</div>
          </div>
          <div class="actions">
            <a class="open" href="{safe_link}" target="_blank" rel="noopener">開啟</a>
            <button class="copy" data-link="{safe_link}">複製連結</button>
          </div>
        </div>
        """

    cards_html += """
    </div>
    <div id="toast" class="toast">已複製連結</div>
    <script>
      const toast = document.getElementById('toast');
      function showToast(){
        toast.classList.add('show');
        setTimeout(()=>toast.classList.remove('show'), 1200);
      }
      document.querySelectorAll('.copy').forEach(btn=>{
        btn.addEventListener('click', async ()=>{
          const link = btn.getAttribute('data-link');
          try{
            await navigator.clipboard.writeText(link);
            showToast();
          }catch(e){
            // fallback
            const ta = document.createElement('textarea');
            ta.value = link;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            document.body.removeChild(ta);
            showToast();
          }
        });
      });
    </script>
    """

    height = 110 + len(items) * 78
    height = min(max(height, 220), 900)
    components.html(cards_html, height=height, scrolling=True)

# ==========================================
#              Core Logic Function
# ==========================================
def execute_automation_logic(bot, source_path, file_prefix, jobs, auto_clean):
    main_progress = st.progress(0, text="準備開始…")
    status_area = st.empty()
    detail_bar_placeholder = st.empty()

    sorted_jobs = sorted(jobs, key=lambda x: x['start'])

    def set_status(kind, text):
        cls = "blue" if kind == "blue" else ("warn" if kind == "warn" else ("err" if kind == "err" else "gray"))
        status_area.markdown(f"<div class='callout {cls}'>{text}</div>", unsafe_allow_html=True)

    def update_step1(filename, current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"影片上傳：{filename}（{int(pct*100)}%）")

    def update_step2(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"投影片處理：{current}/{total}（{int(pct*100)}%）")

    def update_step3(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"內部檔案優化：{current}/{total}（{int(pct*100)}%）")

    def update_step4(filename, current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"發布上傳：{filename}（{int(pct*100)}%）")

    def update_step5(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"播放器優化：{current}/{total}（{int(pct*100)}%）")

    def general_log(msg):
        print(f"[Log] {msg}")

    try:
        # Step 1
        set_status("blue", "步驟 1/5：提取簡報內影片並上傳至雲端")
        main_progress.progress(5, text="步驟 1：影片雲端化")
        video_map = bot.extract_and_upload_videos(
            source_path,
            os.path.join(WORK_DIR, "media"),
            file_prefix=file_prefix,
            progress_callback=update_step1,
            log_callback=general_log
        )
        detail_bar_placeholder.empty()

        # Step 2
        set_status("blue", "步驟 2/5：以雲端連結圖片替換簡報內影片")
        main_progress.progress(25, text="步驟 2：連結置換")
        mod_path = os.path.join(WORK_DIR, "modified.pptx")
        bot.replace_videos_with_images(
            source_path,
            mod_path,
            video_map,
            progress_callback=update_step2
        )
        detail_bar_placeholder.empty()

        # Step 3
        set_status("blue", "步驟 3/5：檔案瘦身與壓縮（維持可用解析度）")
        main_progress.progress(45, text="步驟 3：檔案優化")
        slim_path = os.path.join(WORK_DIR, "slim.pptx")
        bot.shrink_pptx(
            mod_path,
            slim_path,
            progress_callback=update_step3
        )
        detail_bar_placeholder.empty()

        # Step 4
        set_status("blue", "步驟 4/5：依任務設定拆分簡報並發布至 Google Slides")
        main_progress.progress(65, text="步驟 4：拆分發布")
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
            set_status("err", "流程終止：偵測到拆分後的檔案超出 Google 100MB 限制")
            for err_job in oversized_errors:
                st.error(f"任務「{err_job['filename']}」壓縮後仍有 {err_job['size_mb']:.2f} MB，超過限制（100MB）。")
            st.markdown("<div class='callout warn'>建議：縮小頁數範圍或拆成多個任務後重試。</div>", unsafe_allow_html=True)
            return

        # Step 5
        set_status("blue", "步驟 5/5：優化線上簡報的影片播放器")
        main_progress.progress(85, text="步驟 5：內嵌優化")
        final_results = bot.embed_videos_in_slides(
            results,
            progress_callback=update_step5,
            log_callback=general_log
        )
        detail_bar_placeholder.empty()

        # Final log
        set_status("blue", "最後步驟：寫入資料庫（Google Sheets）")
        main_progress.progress(95, text="寫入資料庫")
        bot.log_to_sheets(final_results, log_callback=general_log)

        main_progress.progress(100, text="完成")
        set_status("blue", "流程已完成：所有自動化步驟成功執行")

        if auto_clean:
            cleanup_workspace()
            st.toast("已清除暫存檔案", icon="🧹")

        st.divider()
        render_result_cards(file_prefix, final_results)

        # 一鍵回到第一步
        st.divider()
        if st.button("返回並處理新檔", use_container_width=True):
            reset_to_step1(keep_bot=True)

    except Exception as e:
        set_status("err", f"執行過程中發生錯誤：{e}")
        with st.expander("查看詳細錯誤資訊"):
            st.code(traceback.format_exc())

# ==========================================
#              Main UI
# ==========================================

# Header
st.markdown(f"""
<div class="auro-header">
  <img src="{LOGO_URL}" alt="AUROTEK" />
  <div class="auro-subtitle">簡報案例自動化發布平台</div>
</div>
""", unsafe_allow_html=True)

# 功能說明（統一企業藍 callout）
st.markdown("""
<div class="callout blue">
功能說明：上傳簡報 → 線上拆分 → 影片雲端化 → 內嵌優化 → 雲端發布 → 寫入和椿資料庫
</div>
""", unsafe_allow_html=True)

# 初始化狀態
if 'split_jobs' not in st.session_state:
    st.session_state.split_jobs = []
if 'ppt_meta' not in st.session_state:
    st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}
if 'current_file_name' not in st.session_state:
    st.session_state.current_file_name = None
if 'bot' not in st.session_state:
    try:
        bot_instance = PPTAutomationBot()
        if bot_instance.creds:
            st.session_state.bot = bot_instance
        else:
            st.markdown("<div class='callout warn'>系統未檢測到有效憑證（Secrets），請確認部署環境設定。</div>", unsafe_allow_html=True)
            st.session_state.bot = bot_instance
    except Exception as e:
        st.markdown(f"<div class='callout err'>Bot 初始化失敗：{e}</div>", unsafe_allow_html=True)

# =========================
# Step 1：檔案來源
# =========================
with st.container():
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("步驟一：選擇檔案來源")

    input_method = st.radio("上傳方式", ["本地檔案", "線上檔案"], horizontal=True)

    ensure_workspace()
    source_path = os.path.join(WORK_DIR, "source.pptx")
    file_name_for_logic = None

    if input_method == "本地檔案":
        uploaded_file = st.file_uploader("PPTX", type=['pptx'], label_visibility="collapsed")
        if uploaded_file:
            file_name_for_logic = uploaded_file.name

            # 重要：只有在「換檔」時才清空工作區，避免刪掉剛寫入的 source.pptx
            if st.session_state.current_file_name != file_name_for_logic:
                cleanup_workspace()

            with open(source_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

    else:
        url_input = st.text_input("PPTX 直接下載網址", placeholder="https://example.com/file.pptx")
        if url_input:
            if not url_input.lower().endswith(".pptx"):
                st.markdown("<div class='callout warn'>提醒：網址結尾似乎不是 .pptx，請確認是否為直接下載連結。</div>", unsafe_allow_html=True)

            fake_name = url_input.split("/")[-1].split("?")[0]
            if not fake_name.lower().endswith(".pptx"):
                fake_name += ".pptx"

            if st.button("下載並載入", use_container_width=True):
                with st.spinner("下載中…"):
                    cleanup_workspace()
                    success, error = download_file_from_url(url_input, source_path)
                    if success:
                        file_name_for_logic = fake_name
                        st.toast("下載完成", icon="✅")
                    else:
                        st.markdown(f"<div class='callout err'>下載失敗：{error}</div>", unsafe_allow_html=True)

    # 解析檔案與預覽
    if file_name_for_logic and os.path.exists(source_path):
        if st.session_state.current_file_name != file_name_for_logic:
            # 換檔：載入歷史任務與重新解析
            saved_jobs = load_history(file_name_for_logic)
            st.session_state.split_jobs = saved_jobs if saved_jobs else []

            progress_placeholder = st.empty()
            progress_placeholder.progress(0, text="解析簡報…")

            try:
                prs = Presentation(source_path)
                total_slides = len(prs.slides)

                preview_data = []
                for i, slide in enumerate(prs.slides):
                    txt = slide.shapes.title.text if (slide.shapes.title and slide.shapes.title.text) else "無標題"
                    if txt == "無標題":
                        for s in slide.shapes:
                            if hasattr(s, "text") and s.text.strip():
                                txt = s.text.strip()[:20] + "…"
                                break
                    preview_data.append({"頁碼": i + 1, "內容摘要": txt})

                st.session_state.ppt_meta["total_slides"] = total_slides
                st.session_state.ppt_meta["preview_data"] = preview_data
                st.session_state.current_file_name = file_name_for_logic

                progress_placeholder.progress(100, text="完成")
                st.markdown(
                    f"<div class='callout blue'>已讀取：{file_name_for_logic}（共 {total_slides} 頁）</div>",
                    unsafe_allow_html=True
                )

            except Exception as e:
                st.markdown(f"<div class='callout err'>檔案處理失敗：{e}</div>", unsafe_allow_html=True)
                st.session_state.current_file_name = None
                st.stop()

    st.markdown("</div>", unsafe_allow_html=True)

# =========================
# Step 2：拆分任務
# =========================
if st.session_state.current_file_name:
    total_slides = st.session_state.ppt_meta["total_slides"]
    preview_data = st.session_state.ppt_meta["preview_data"]

    with st.expander("頁碼與標題對照表", expanded=False):
        st.dataframe(preview_data, use_container_width=True, height=260, hide_index=True)

    with st.container():
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        col_a, col_b = st.columns([3, 1])
        col_a.subheader("步驟二：設定拆分任務")
        if col_b.button("新增任務", type="primary", use_container_width=True):
            add_split_job(total_slides)

        if not st.session_state.split_jobs:
            st.markdown("<div class='callout gray'>尚未建立任務，請先新增任務並設定頁數範圍。</div>", unsafe_allow_html=True)

        for i, job in enumerate(st.session_state.split_jobs):
            with st.container(border=True):
                st.markdown(f"**任務 {i+1}**")

                c1, c2, c3 = st.columns([3, 1.5, 1.5])
                job["filename"] = c1.text_input("檔名", value=job["filename"], key=f"f_{job['id']}", placeholder="例如：清潔案例A")
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

        # 保存歷史任務
        save_history(st.session_state.current_file_name, st.session_state.split_jobs)

        st.markdown("</div>", unsafe_allow_html=True)

# =========================
# Step 3：執行
# =========================
if st.session_state.current_file_name:
    with st.container():
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("步驟三：開始執行")

        auto_clean = st.checkbox("任務完成後自動清除暫存檔", value=True)

        if st.button("執行自動化排程", type="primary", use_container_width=True):
            if not st.session_state.split_jobs:
                st.markdown("<div class='callout err'>請至少設定一個拆分任務後再執行。</div>", unsafe_allow_html=True)
            else:
                validation_errors = validate_jobs(st.session_state.split_jobs, st.session_state.ppt_meta["total_slides"])
                if validation_errors:
                    for err in validation_errors:
                        st.error(err)
                    st.markdown("<div class='callout err'>請修正上述錯誤後再執行。</div>", unsafe_allow_html=True)
                else:
                    if 'bot' not in st.session_state or not st.session_state.bot:
                        st.markdown("<div class='callout err'>機器人未初始化（憑證錯誤），請檢查 Secrets。</div>", unsafe_allow_html=True)
                        st.stop()

                    execute_automation_logic(
                        st.session_state.bot,
                        os.path.join(WORK_DIR, "source.pptx"),
                        os.path.splitext(st.session_state.current_file_name)[0],
                        st.session_state.split_jobs,
                        auto_clean
                    )

        st.markdown("</div>", unsafe_allow_html=True)
