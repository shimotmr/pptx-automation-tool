import streamlit as st
import streamlit.components.v1 as components
import os
import uuid
import json
import shutil
import traceback
import requests
import hashlib
from datetime import datetime
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
MANIFEST_FILE = "processed_manifest.json"

# ==========================================
#              企業版 CSS（保留功能、重做風格）
# ==========================================
st.markdown("""
<style>
header[data-testid="stHeader"] { display: none; }
.stApp > header { display: none; }

.block-container {
  padding-top: 0.9rem !important;
  padding-bottom: 1.0rem !important;
}

h3 { font-size: 1.35rem !important; font-weight: 700 !important; }
h4 { font-size: 1.05rem !important; font-weight: 650 !important; color: #1f2937; }
[data-testid="stAlert"] p { font-size: 0.90rem !important; line-height: 1.45 !important; }

:root{
  --brand-blue:#0B4F8A;
  --brand-blue-weak:#EAF3FF;
  --border:#E5E7EB;
  --text:#111827;
  --muted:#6B7280;
  --bg-soft:#F8FAFC;
}

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

.section-card{
  border:1px solid var(--border);
  border-radius:16px;
  padding: 14px 14px 6px 14px;
  background:#fff;
}

.stProgress > div > div > div > div { color: white; font-weight: 600; }

/* ==========================================
   FileUploader：修正「瀏覽檔案」重複 / 縱排 / 框線錯位
========================================== */
[data-testid="stFileUploaderDropzoneInstructions"] > div:first-child { display:none !important; }
[data-testid="stFileUploaderDropzoneInstructions"] > div:nth-child(2) { display:none !important; }

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

section[data-testid="stFileUploaderDropzone"]{
  padding: 0.60rem 0.90rem !important;
  border-radius:14px !important;
  background: var(--bg-soft) !important;
}

section[data-testid="stFileUploaderDropzone"] button{
  font-size:0 !important;
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

div[data-testid="stFileUploader"] section:not([data-testid="stFileUploaderDropzone"]) button{
  display:none !important;
}

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
    if os.path.exists(WORK_DIR):
        try:
            shutil.rmtree(WORK_DIR)
        except Exception as e:
            print(f"Cleanup warning: {e}")
    os.makedirs(WORK_DIR, exist_ok=True)

def sha256_of_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return default
    return default

def save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Save json failed: {e}")

def load_history(filename):
    data = load_json(HISTORY_FILE, {})
    return data.get(filename, [])

def save_history(filename, jobs):
    data = load_json(HISTORY_FILE, {})
    data[filename] = jobs
    save_json(HISTORY_FILE, data)

def load_manifest():
    return load_json(MANIFEST_FILE, {})

def save_manifest(m):
    save_json(MANIFEST_FILE, m)

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
        cur = sorted_jobs[i]
        nxt = sorted_jobs[i+1]
        if cur['end'] >= nxt['start']:
            errors.append(
                f"⚠️ 頁數重疊：{cur['filename']}（{cur['start']}-{cur['end']}）與 {nxt['filename']}（{nxt['start']}-{nxt['end']}）"
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

def scroll_to_anchor(anchor_id: str):
    components.html(
        f"""
        <script>
          const el = window.parent.document.getElementById("{anchor_id}");
          if(el) {{
            el.scrollIntoView({{behavior:"smooth", block:"start"}});
          }}
        </script>
        """,
        height=0
    )

def reset_to_step1(keep_bot=True):
    # 讓 uploader widget 重新初始化，避免「回到 step3」
    st.session_state.uploader_key = str(uuid.uuid4())[:8]

    keys = [
        "current_file_name", "ppt_meta", "split_jobs", "total_slides",
        "source_hash", "source_prefix", "force_rerun", "prefix_override"
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

    # JS：複製後顯示「藍色提示卡」，1.2 秒後消失
    cards_html = """
    <style>
      .wrap{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Noto Sans TC","PingFang TC",Arial;}
      .card{
        border:1px solid #E5E7EB;border-radius:14px;padding:12px 14px;margin:10px 0;
        display:flex;align-items:center;justify-content:space-between;background:#fff;
      }
      .left{display:flex;flex-direction:column;gap:4px;}
      .title{font-weight:750;color:#111827;font-size:14px;}
      .meta{font-size:12px;color:#6B7280;}
      .actions{display:flex;align-items:center;gap:10px;flex-wrap:nowrap;}
      .open{
        text-decoration:none;background:#EAF3FF;color:#0B4F8A;padding:8px 10px;border-radius:10px;
        font-weight:750;font-size:13px;border:1px solid #D6E8FF;white-space:nowrap;
      }
      .copy{
        border:1px solid #E5E7EB;background:#F8FAFC;border-radius:10px;padding:8px 10px;
        cursor:pointer;font-weight:750;font-size:13px;white-space:nowrap;
      }

      /* 企業藍提示卡（模仿上方流程完成圖卡） */
      .toastcard{
        position:fixed;
        right:16px;
        bottom:16px;
        width:min(420px, 92vw);
        border:1px solid #D6E8FF;
        border-left:4px solid #0B4F8A;
        background:#EAF3FF;
        color:#0B4F8A;
        padding:12px 14px;
        border-radius:14px;
        font-weight:750;
        opacity:0;
        transform:translateY(8px);
        transition:all .18s ease;
        z-index:9999;
        box-shadow: 0 8px 22px rgba(15,23,42,.08);
      }
      .toastcard.show{opacity:1;transform:translateY(0);}
      .toastrow{display:flex;align-items:center;gap:10px;}
      .dot{
        width:10px;height:10px;border-radius:999px;background:#0B4F8A;flex:0 0 auto;
      }
      .tmsg{font-size:13px;line-height:1.35;}
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

    <div id="toastcard" class="toastcard">
      <div class="toastrow">
        <div class="dot"></div>
        <div class="tmsg">已複製連結到剪貼簿</div>
      </div>
    </div>

    <script>
      const toast = document.getElementById('toastcard');
      let timer = null;

      function showToast(){
        toast.classList.add('show');
        if(timer) clearTimeout(timer);
        timer = setTimeout(()=>toast.classList.remove('show'), 1200);
      }

      document.querySelectorAll('.copy').forEach(btn=>{
        btn.addEventListener('click', async ()=>{
          const link = btn.getAttribute('data-link');
          try{
            await navigator.clipboard.writeText(link);
            showToast();
          }catch(e){
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
def execute_automation_logic(bot, source_path, file_prefix, jobs, auto_clean, source_hash):
    # 自動捲動：進到進度區
    scroll_to_anchor("run-anchor")

    main_progress = st.progress(0, text="準備開始…")
    status_area = st.empty()
    detail_bar_placeholder = st.empty()

    sorted_jobs = sorted(jobs, key=lambda x: x['start'])

    def set_status(kind, text):
        cls = "blue" if kind == "blue" else ("warn" if kind == "warn" else ("err" if kind == "err" else "gray"))
        status_area.markdown(f"<div class='callout {cls}'>{text}</div>", unsafe_allow_html=True)
        # 每次更新狀態都嘗試把視窗維持在進度區附近
        scroll_to_anchor("run-anchor")

    def update_step1(filename, current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"影片上傳：{filename}（{int(pct*100)}%）")

    def update_step2(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"投影片處理：{current}/{total}（{int(pct*100)}%）")

    def update_step3(current, total):
        pct = current / total if total > 0 else 0
        detail_bar_placeholder.progress(pct, text=f"檔案優化：{current}/{total}（{int(pct*100)}%）")

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

        # 寫入 manifest：用 hash 防止重複執行
        manifest = load_manifest()
        manifest[source_hash] = {
            "file_prefix": file_prefix,
            "finished_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "results": [
                {"filename": r.get("filename"), "final_link": r.get("final_link")}
                for r in (final_results or [])
                if r.get("final_link")
            ],
        }
        save_manifest(manifest)

        if auto_clean:
            cleanup_workspace()
            st.toast("已清除暫存檔案", icon="🧹")

        st.divider()
        render_result_cards(file_prefix, final_results)

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

st.markdown("""
<div class="callout blue">
功能說明：上傳簡報 → 線上拆分 → 影片雲端化 → 內嵌優化 → 雲端發布 → 寫入和椿資料庫
</div>
""", unsafe_allow_html=True)

# 初始化狀態
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = str(uuid.uuid4())[:8]
if 'split_jobs' not in st.session_state:
    st.session_state.split_jobs = []
if 'ppt_meta' not in st.session_state:
    st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}
if 'current_file_name' not in st.session_state:
    st.session_state.current_file_name = None
if 'bot' not in st.session_state:
    try:
        bot_instance = PPTAutomationBot()
        st.session_state.bot = bot_instance
        if not getattr(bot_instance, "creds", None):
            st.markdown("<div class='callout warn'>系統未檢測到有效憑證（Secrets），請確認部署環境設定。</div>", unsafe_allow_html=True)
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
        uploaded_file = st.file_uploader(
            "PPTX", type=['pptx'], label_visibility="collapsed",
            key=f"uploader_{st.session_state.uploader_key}"
        )
        if uploaded_file:
            file_name_for_logic = uploaded_file.name

            # 換檔才清空 workspace
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
        # 計算 hash（用於防重複執行）
        source_hash = sha256_of_file(source_path)
        st.session_state.source_hash = source_hash

        if st.session_state.current_file_name != file_name_for_logic:
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

        # 防重複執行（以 hash 為準）
        manifest = load_manifest()
        source_hash = st.session_state.get("source_hash")
        already_done = bool(source_hash and source_hash in manifest)

        # 預設 prefix = 檔名（不含 .pptx）
        default_prefix = os.path.splitext(st.session_state.current_file_name)[0]
        st.session_state.source_prefix = default_prefix

        if already_done:
            info = manifest.get(source_hash, {})
            prev_at = info.get("finished_at", "（未知時間）")
            prev_prefix = info.get("file_prefix", default_prefix)
            st.markdown(
                f"<div class='callout warn'>偵測到此檔案已執行過（{prev_at}），預設將避免重複執行。</div>",
                unsafe_allow_html=True
            )
            st.caption(f"上次使用的輸出前綴：{prev_prefix}")

        force_rerun = False
        prefix_override = default_prefix

        if already_done:
            force_rerun = st.checkbox("仍要重新執行（可能會產生重複雲端結果）", value=False)
            if force_rerun:
                prefix_override = st.text_input(
                    "輸出前綴（建議改名避免混淆）",
                    value=f"{default_prefix}_rerun",
                    help="此名稱會用於雲端資料夾/檔名的前綴，用來區分不同批次"
                )

        # 進度區 anchor（用於自動捲動）
        st.markdown("<div id='run-anchor'></div>", unsafe_allow_html=True)

        run_btn_disabled = already_done and (not force_rerun)

        if st.button("執行自動化排程", type="primary", use_container_width=True, disabled=run_btn_disabled):
            # 點下按鈕立即捲動到進度區
            scroll_to_anchor("run-anchor")

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

                    used_prefix = prefix_override if (already_done and force_rerun) else default_prefix

                    execute_automation_logic(
                        st.session_state.bot,
                        os.path.join(WORK_DIR, "source.pptx"),
                        used_prefix,
                        st.session_state.split_jobs,
                        auto_clean,
                        source_hash
                    )

        if run_btn_disabled:
            st.caption("如需再次執行，請先勾選「仍要重新執行」並建議修改輸出前綴。")

        st.markdown("</div>", unsafe_allow_html=True)
