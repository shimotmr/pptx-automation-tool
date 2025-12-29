import os
import re
import json
import uuid
import time
import shutil
import hashlib
import traceback
import requests

import streamlit as st
import streamlit.components.v1 as components
from pptx import Presentation

from ppt_processor import PPTAutomationBot

# =========================================================
#                     Page Config
# =========================================================
st.set_page_config(
    page_title="Aurotek｜簡報案例自動化發布平台",
    page_icon="📊",
    layout="wide",
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"            # per filename: split jobs
STATE_DB_FILE = "resume_state_db.json"       # reserved (optional)

# =========================================================
#                 Enterprise CSS (Stable)
# =========================================================
st.markdown(
    """
<style>
/* Hide default streamlit header */
header[data-testid="stHeader"] { display: none; }
.stApp > header { display: none; }

/* Layout paddings */
.block-container{
  padding-top: 0.85rem !important;
  padding-bottom: 1.0rem !important;
  max-width: 1200px;
}

/* Typography */
h2, h3 { font-weight: 800 !important; letter-spacing: .2px; }
h4 { font-weight: 750 !important; color:#111827; }
[data-testid="stAlert"] p { font-size: 0.92rem !important; line-height: 1.5 !important; }

/* Brand */
:root{
  --brand-blue:#0B4F8A;
  --brand-blue-weak:#EAF3FF;
  --border:#E5E7EB;
  --text:#111827;
  --muted:#6B7280;
  --bg-soft:#F8FAFC;
  --bg:#FFFFFF;
}

/* Header */
.auro-header{
  display:flex;
  flex-direction:column;
  align-items:center;
  justify-content:center;
  margin: 0 0 10px 0;
}
.auro-header img{
  width:300px; /* desktop lock */
  height:auto;
  max-width: 90vw;
}
.auro-subtitle{
  margin-top:4px;
  color: var(--muted);
  font-size: 1.00rem;
  font-weight: 650;
  letter-spacing: 1.6px;
  text-align:center;
}

/* Callouts */
.callout{
  border:1px solid var(--border);
  border-radius:14px;
  padding:12px 14px;
  margin: 10px 0;
  background: var(--bg);
}
.callout.blue{
  border-left: 4px solid var(--brand-blue);
  background: var(--brand-blue-weak);
  color: var(--brand-blue);
  font-weight: 750;
}
.callout.gray{
  background: var(--bg-soft);
  color: var(--text);
}
.callout.warn{
  border-left: 4px solid #B45309;
  background:#FFF7ED;
  color:#92400E;
  font-weight:750;
}
.callout.err{
  border-left: 4px solid #B91C1C;
  background:#FEF2F2;
  color:#991B1B;
  font-weight:750;
}

/* Section card */
.section-card{
  border:1px solid var(--border);
  border-radius:18px;
  padding: 14px 14px 8px 14px;
  background: var(--bg);
  margin-bottom: 14px;
}

/* Progress bar text style */
.stProgress > div > div > div > div { color: white; font-weight: 700; }

/* ===========================
   FileUploader Fix:
   1) No vertical duplicated Browse
   2) Only one Browse button (dropzone)
   3) Avoid text-over-border bug
=========================== */
[data-testid="stFileUploaderDropzoneInstructions"] > div:first-child { display:none !important; }
[data-testid="stFileUploaderDropzoneInstructions"] > div:nth-child(2) { display:none !important; }

[data-testid="stFileUploaderDropzoneInstructions"]::before{
  content:"拖放或點擊上傳";
  display:block;
  font-size:0.95rem;
  font-weight:800;
  line-height:1.2;
  margin:0;
}
[data-testid="stFileUploaderDropzoneInstructions"]::after{
  content:"PPTX · 單檔 5GB";
  display:block;
  font-size:0.76rem;
  color: var(--muted);
  line-height:1.15;
  margin-top:2px;
}

section[data-testid="stFileUploaderDropzone"]{
  padding: 0.62rem 0.95rem !important;
  border-radius:16px !important;
  background: var(--bg-soft) !important;
}

section[data-testid="stFileUploaderDropzone"] button{
  font-size:0 !important;  /* hide original text safely */
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
  font-weight:800;
  color:#111827;
}

div[data-testid="stFileUploader"] section:not([data-testid="stFileUploaderDropzone"]) button{
  display:none !important; /* hide duplicate button in file list */
}

/* Mobile logo */
@media (max-width: 768px){
  .block-container { padding-top:0.7rem !important; }
  .auro-header img { width: 260px; }
  .auro-subtitle { font-size:0.98rem; letter-spacing:1px; }
}
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
#                     Utility Helpers
# =========================================================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)


def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        try:
            shutil.rmtree(WORK_DIR)
        except Exception:
            pass
    os.makedirs(WORK_DIR, exist_ok=True)


def _safe_filename_key(name: str) -> str:
    name = name or "unknown"
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", name)


def _read_json(path, default):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        return default
    return default


def _write_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def load_history(filename):
    db = _read_json(HISTORY_FILE, {})
    return db.get(filename, [])


def save_history(filename, jobs):
    db = _read_json(HISTORY_FILE, {})
    db[filename] = jobs
    _write_json(HISTORY_FILE, db)


def file_fingerprint(path: str) -> str:
    """Fast-ish fingerprint: size + mtime + first/last 1MB sha1 (won't read huge file fully)."""
    try:
        st_ = os.stat(path)
        size = st_.st_size
        mtime = int(st_.st_mtime)
        h = hashlib.sha1()
        h.update(f"{size}:{mtime}".encode("utf-8"))
        with open(path, "rb") as f:
            head = f.read(1024 * 1024)
            if head:
                h.update(head)
            if size > 1024 * 1024:
                try:
                    f.seek(max(0, size - 1024 * 1024))
                    tail = f.read(1024 * 1024)
                    if tail:
                        h.update(tail)
                except Exception:
                    pass
        return h.hexdigest()
    except Exception:
        return str(uuid.uuid4())


def scroll_to_bottom():
    components.html(
        """
<script>
  try{
    const doc = window.parent.document;
    const main = doc.querySelector('section.main');
    if(main){ main.scrollTo({top: main.scrollHeight, behavior:'smooth'}); }
    else { window.parent.scrollTo({top: doc.body.scrollHeight, behavior:'smooth'}); }
  }catch(e){}
</script>
""",
        height=0,
    )


def validate_jobs(jobs, total_slides):
    errors = []
    for i, job in enumerate(jobs):
        label = f"任務 {i+1}（{job.get('filename') or '未命名'}）"
        if not str(job.get("filename", "")).strip():
            errors.append(f"❌ {label}：檔名不能為空。")
        if int(job.get("start", 1)) > int(job.get("end", 1)):
            errors.append(f"❌ {label}：起始頁不能大於結束頁。")
        if int(job.get("end", 1)) > int(total_slides):
            errors.append(f"❌ {label}：結束頁超出簡報總頁數（{total_slides}）。")

    sorted_jobs = sorted(jobs, key=lambda x: int(x.get("start", 1)))
    for i in range(len(sorted_jobs) - 1):
        a = sorted_jobs[i]
        b = sorted_jobs[i + 1]
        if int(a.get("end", 1)) >= int(b.get("start", 1)):
            errors.append(
                f"⚠️ 頁數重疊：{a.get('filename','')}（{a.get('start')}-{a.get('end')}）"
                f" 與 {b.get('filename','')}（{b.get('start')}-{b.get('end')}）"
            )
    return errors


def download_file_from_url(url, dest_path):
    try:
        r = requests.get(url, stream=True, timeout=60)
        r.raise_for_status()
        with open(dest_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        return True, None
    except Exception as e:
        return False, str(e)


def reset_to_step1(keep_bot=True):
    for k in [
        "current_file_name",
        "ppt_meta",
        "split_jobs",
        "total_slides",
    ]:
        if k in st.session_state:
            del st.session_state[k]
    if (not keep_bot) and ("bot" in st.session_state):
        del st.session_state["bot"]
    cleanup_workspace()
    st.rerun()


# =========================================================
#     Google Sheet Category Options (from Presentations!B:B)
# =========================================================
def get_category_options(bot: PPTAutomationBot):
    fallback = ["清潔", "配送", "購物", "AURO"]
    try:
        if not getattr(bot, "sheets_service", None):
            return fallback

        # Try bot.SPREADSHEET_ID if exists
        spreadsheet_id = getattr(bot, "SPREADSHEET_ID", None)
        if not spreadsheet_id:
            try:
                from ppt_processor import SPREADSHEET_ID  # type: ignore
                spreadsheet_id = SPREADSHEET_ID
            except Exception:
                spreadsheet_id = None

        if not spreadsheet_id:
            return fallback

        resp = bot.sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range="Presentations!B:B",
        ).execute()

        vals = resp.get("values", [])
        cats = []
        for row in vals:
            if row and str(row[0]).strip():
                cats.append(str(row[0]).strip())

        seen, out = set(), []
        for c in cats:
            if c not in seen:
                seen.add(c)
                out.append(c)

        return out or fallback
    except Exception:
        return fallback


# =========================================================
#           Split Jobs Helpers (UI + Persistence)
# =========================================================
def new_job_template(total_pages, categories):
    # ✅ 修正：避免 total_pages=0 導致 number_input value 不合法 → 當機
    total_pages = max(1, int(total_pages or 0))
    return {
        "id": str(uuid.uuid4())[:8],
        "filename": "",
        "start": 1,
        "end": total_pages,
        "category": categories[0] if categories else "清潔",
        "subcategory": "",
        "client": "",
        "keywords": "",
    }


def add_split_job(total_pages, categories):
    st.session_state.split_jobs.insert(0, new_job_template(total_pages, categories))


def remove_split_job(index):
    st.session_state.split_jobs.pop(index)


# =========================================================
#           Resume state paths (per filename, local)
# =========================================================
def state_paths(file_name: str):
    k = _safe_filename_key(file_name)
    return {
        "state_json": os.path.join(WORK_DIR, f"state_{k}.json"),
        "video_map": os.path.join(WORK_DIR, f"video_map_{k}.json"),
        "mod_pptx": os.path.join(WORK_DIR, f"modified_{k}.pptx"),
        "slim_pptx": os.path.join(WORK_DIR, f"slim_{k}.pptx"),
        "split_json": os.path.join(WORK_DIR, f"split_results_{k}.json"),
        "final_json": os.path.join(WORK_DIR, f"final_results_{k}.json"),
        "fingerprint_txt": os.path.join(WORK_DIR, f"fingerprint_{k}.txt"),
    }


def load_pipeline_state(file_name: str):
    return _read_json(state_paths(file_name)["state_json"], {})


def save_pipeline_state(file_name: str, state: dict):
    _write_json(state_paths(file_name)["state_json"], state)


def mark_stage(file_name: str, stage: str, extra: dict | None = None):
    state = load_pipeline_state(file_name)
    state["stage"] = stage
    state["updated_at"] = int(time.time())
    if extra:
        state.update(extra)
    save_pipeline_state(file_name, state)


def detect_resume(file_name: str, source_fp: str):
    """Return (can_resume, info_dict)."""
    paths = state_paths(file_name)
    state = load_pipeline_state(file_name)

    saved_fp = None
    try:
        if os.path.exists(paths["fingerprint_txt"]):
            saved_fp = open(paths["fingerprint_txt"], "r", encoding="utf-8").read().strip()
    except Exception:
        saved_fp = None

    if (not state) or (not saved_fp) or (saved_fp != source_fp):
        return False, {"reason": "no_match"}

    stage = state.get("stage", "none")

    ok = False
    if stage == "videos_done":
        ok = os.path.exists(paths["video_map"])
    elif stage == "replace_done":
        ok = os.path.exists(paths["mod_pptx"])
    elif stage == "shrink_done":
        ok = os.path.exists(paths["slim_pptx"])
    elif stage == "split_done":
        ok = os.path.exists(paths["split_json"])
    elif stage in ("embed_done", "logged_done"):
        ok = os.path.exists(paths["final_json"])
    else:
        ok = False

    return ok, {"stage": stage, "paths": paths, "state": state}


# =========================================================
#            Safe helpers for video-less decks
# =========================================================
def safe_replace_videos(bot: PPTAutomationBot, source_path: str, mod_path: str, video_map: dict, progress_callback=None):
    """
    - If no videos detected => copy source -> mod_path
    - If videos exist => bot.replace_videos_with_images(...)
    """
    video_map = video_map or {}
    if len(video_map) == 0:
        shutil.copyfile(source_path, mod_path)
        if progress_callback:
            try:
                progress_callback(1, 1)
            except Exception:
                pass
        return mod_path

    return bot.replace_videos_with_images(
        source_path,
        mod_path,
        video_map,
        progress_callback=progress_callback,
    )


# =========================================================
#            Results UI (buttons + copy + list)
# =========================================================
def render_completion_card(file_prefix: str, final_results: list[dict]):
    items = []
    for r in final_results or []:
        link = r.get("final_link")
        if link:
            items.append((r.get("filename", ""), link))

    if not items:
        st.markdown("<div class='callout warn'>未產生任何結果連結，請檢查是否有任務被跳過。</div>", unsafe_allow_html=True)
        return

    st.subheader("完成上傳與寫入資料庫")
    st.markdown(
        "<div class='callout blue'>已完成所有步驟，以下為本次產出的線上簡報連結（可開啟 / 一鍵複製）。</div>",
        unsafe_allow_html=True,
    )

    cards_html = """
    <style>
      .wrap{font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Noto Sans TC","PingFang TC",Arial;}
      .card{
        border:1px solid #E5E7EB;
        border-radius:16px;
        padding:12px 14px;
        margin:10px 0;
        display:flex;
        align-items:center;
        justify-content:space-between;
        background:#fff;
        gap:12px;
      }
      .left{display:flex;flex-direction:column;gap:4px;min-width:0;}
      .title{font-weight:800;color:#111827;font-size:14px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width: 520px;}
      .meta{font-size:12px;color:#6B7280;}
      .actions{display:flex;align-items:center;gap:10px;flex-wrap:nowrap;}
      .open{
        text-decoration:none;
        background:#EAF3FF;
        color:#0B4F8A;
        padding:8px 10px;
        border-radius:12px;
        font-weight:800;
        font-size:13px;
        border:1px solid #D6E8FF;
        white-space:nowrap;
      }
      .copy{
        border:1px solid #E5E7EB;
        background:#F8FAFC;
        border-radius:12px;
        padding:8px 10px;
        cursor:pointer;
        font-weight:800;
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
        font-weight:800;
        font-size:13px;
        opacity:0;
        transform: translateY(6px);
        transition: all .18s ease;
        z-index:9999;
      }
      .toast.show{ opacity:1; transform: translateY(0px); }
    </style>
    <div class="wrap">
    """

    for name, link in items:
        display = f"[{file_prefix}]_{name}" if name else f"[{file_prefix}]"
        safe_display = display.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        safe_link = link.replace('"', "%22")
        cards_html += f"""
        <div class="card">
          <div class="left">
            <div class="title">{safe_display}</div>
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
        setTimeout(()=>toast.classList.remove('show'), 1300);
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

    height = min(max(240, 130 + len(items) * 78), 860)
    components.html(cards_html, height=height, scrolling=True)


# =========================================================
#                 Core Execution (Resume-ready)
# =========================================================
def execute_automation_logic(bot: PPTAutomationBot, source_path: str, file_name: str, file_prefix: str, jobs: list[dict], auto_clean: bool):
    ensure_workspace()
    paths = state_paths(file_name)

    status_area = st.empty()
    detail_bar = st.empty()
    main_bar = st.progress(0, text="準備開始…")

    def set_status(kind: str, text: str):
        cls = "blue"
        if kind == "warn":
            cls = "warn"
        elif kind == "err":
            cls = "err"
        elif kind == "gray":
            cls = "gray"
        status_area.markdown(f"<div class='callout {cls}'>{text}</div>", unsafe_allow_html=True)
        scroll_to_bottom()

    def pct_text(step_name: str, pct: float):
        pct = max(0.0, min(1.0, pct))
        return f"{step_name} {int(pct*100)}%"

    def update_step_video(filename, current, total):
        pct = (current / total) if total else 0
        detail_bar.progress(pct, text=pct_text(f"影片上傳：{filename}", pct))

    def update_step_replace(current, total):
        pct = (current / total) if total else 0
        detail_bar.progress(pct, text=pct_text(f"影片置換：{current}/{total}", pct))

    def update_step_shrink(current, total):
        pct = (current / total) if total else 0
        detail_bar.progress(pct, text=pct_text(f"檔案優化：{current}/{total}", pct))

    def update_step_split(filename, current, total):
        pct = (current / total) if total else 0
        detail_bar.progress(pct, text=pct_text(f"拆分上傳：{filename}", pct))

    def update_step_embed(current, total):
        pct = (current / total) if total else 0
        detail_bar.progress(pct, text=pct_text(f"內嵌優化：{current}/{total}", pct))

    def log_cb(msg: str):
        print(f"[APP] {msg}")

    src_fp = file_fingerprint(source_path)
    try:
        with open(paths["fingerprint_txt"], "w", encoding="utf-8") as f:
            f.write(src_fp)
    except Exception:
        pass

    can_resume, info = detect_resume(file_name, src_fp)
    state = info.get("state", {}) if isinstance(info, dict) else {}
    stage = state.get("stage", "none") if can_resume else "none"

    start_from = "videos"
    if can_resume:
        if stage == "videos_done":
            start_from = "replace"
        elif stage == "replace_done":
            start_from = "shrink"
        elif stage == "shrink_done":
            start_from = "split"
        elif stage == "split_done":
            start_from = "embed"
        elif stage == "embed_done":
            start_from = "log"
        elif stage == "logged_done":
            start_from = "done"
        else:
            start_from = "videos"
        set_status("blue", f"偵測到可從斷點繼續（已完成：{stage}），將從下一步接續執行。")

    sorted_jobs = sorted(jobs, key=lambda x: int(x.get("start", 1)))

    try:
        scroll_to_bottom()

        # Step 1
        video_map = {}
        if start_from == "videos":
            set_status("blue", "步驟 1/5：提取簡報內影片並上傳至雲端")
            main_bar.progress(0.06, text="步驟 1/5：影片雲端化 0%")
            video_map = bot.extract_and_upload_videos(
                source_path,
                os.path.join(WORK_DIR, "media"),
                file_prefix=file_prefix,
                progress_callback=update_step_video,
                log_callback=log_cb,
            ) or {}
            _write_json(paths["video_map"], video_map)
            mark_stage(file_name, "videos_done", {"video_map_count": len(video_map)})
            detail_bar.empty()
            main_bar.progress(0.22, text="步驟 1/5：影片雲端化 100%")
        else:
            video_map = _read_json(paths["video_map"], {})
            if not isinstance(video_map, dict):
                video_map = {}

        # Step 2
        if start_from in ("videos", "replace"):
            set_status("blue", "步驟 2/5：以雲端連結圖片替換簡報內影片（若無影片則自動跳過）")
            main_bar.progress(0.26, text="步驟 2/5：影片置換 0%")
            safe_replace_videos(
                bot,
                source_path,
                paths["mod_pptx"],
                video_map,
                progress_callback=update_step_replace,
            )
            if not os.path.exists(paths["mod_pptx"]):
                raise FileNotFoundError(f"replace 產物不存在：{paths['mod_pptx']}")
            mark_stage(file_name, "replace_done")
            detail_bar.empty()
            main_bar.progress(0.42, text="步驟 2/5：影片置換 100%")

        # Step 3
        if start_from in ("videos", "replace", "shrink"):
            set_status("blue", "步驟 3/5：檔案瘦身與壓縮（維持可用解析度）")
            main_bar.progress(0.46, text="步驟 3/5：檔案優化 0%")

            shrink_in = paths["mod_pptx"] if os.path.exists(paths["mod_pptx"]) else source_path
            if not os.path.exists(shrink_in):
                raise FileNotFoundError(f"找不到 shrink 輸入檔：{shrink_in}")

            bot.shrink_pptx(
                shrink_in,
                paths["slim_pptx"],
                progress_callback=update_step_shrink,
            )
            if not os.path.exists(paths["slim_pptx"]):
                raise FileNotFoundError(f"shrink 產物不存在：{paths['slim_pptx']}")
            mark_stage(file_name, "shrink_done")
            detail_bar.empty()
            main_bar.progress(0.62, text="步驟 3/5：檔案優化 100%")

        # Step 4
        split_results = []
        if start_from in ("videos", "replace", "shrink", "split"):
            set_status("blue", "步驟 4/5：依任務設定拆分簡報並發布至 Google Slides")
            main_bar.progress(0.66, text="步驟 4/5：拆分發布 0%")

            split_in = paths["slim_pptx"] if os.path.exists(paths["slim_pptx"]) else (
                paths["mod_pptx"] if os.path.exists(paths["mod_pptx"]) else source_path
            )
            if not os.path.exists(split_in):
                raise FileNotFoundError(f"找不到 split 輸入檔：{split_in}")

            split_results = bot.split_and_upload(
                split_in,
                sorted_jobs,
                file_prefix=file_prefix,
                progress_callback=update_step_split,
                log_callback=log_cb,
            ) or []

            _write_json(paths["split_json"], split_results)
            mark_stage(file_name, "split_done")
            detail_bar.empty()
            main_bar.progress(0.82, text="步驟 4/5：拆分發布 100%")
        else:
            split_results = _read_json(paths["split_json"], [])
            if not isinstance(split_results, list):
                split_results = []

        oversized = [r for r in split_results if r.get("error_too_large")]
        if oversized:
            set_status("err", "流程終止：偵測到拆分後的檔案超出 Google 100MB 限制")
            for r in oversized:
                st.error(f"任務「{r.get('filename','')}」壓縮後仍有 {float(r.get('size_mb',0)):.2f} MB，超過限制（100MB）。")
            st.markdown("<div class='callout warn'>建議：縮小頁數範圍或拆成多個任務後重試。</div>", unsafe_allow_html=True)
            return

        # Step 5
        final_results = []
        if start_from in ("videos", "replace", "shrink", "split", "embed"):
            set_status("blue", "步驟 5/5：優化線上簡報的影片播放器（若無影片則快速跳過）")
            main_bar.progress(0.86, text="步驟 5/5：內嵌優化 0%")

            final_results = bot.embed_videos_in_slides(
                split_results,
                progress_callback=update_step_embed,
                log_callback=log_cb,
            ) or []

            _write_json(paths["final_json"], final_results)
            mark_stage(file_name, "embed_done")
            detail_bar.empty()
            main_bar.progress(0.94, text="步驟 5/5：內嵌優化 100%")
        else:
            final_results = _read_json(paths["final_json"], [])
            if not isinstance(final_results, list):
                final_results = []

        # Log to Sheets
        if start_from in ("videos", "replace", "shrink", "split", "embed", "log"):
            set_status("blue", "最後步驟：寫入資料庫（Google Sheets）")
            main_bar.progress(0.97, text="寫入資料庫 0%")
            bot.log_to_sheets(final_results, log_callback=log_cb)
            mark_stage(file_name, "logged_done")
            main_bar.progress(1.0, text="完成 100%")
        else:
            main_bar.progress(1.0, text="完成 100%")

        set_status("blue", "流程已完成：已完成上傳、內嵌優化並寫入資料庫")

        st.divider()
        render_completion_card(file_prefix, final_results)

        st.divider()
        colx, coly = st.columns([1, 1])
        with colx:
            if st.button("返回並處理新檔", use_container_width=True):
                reset_to_step1(keep_bot=True)
        with coly:
            if st.button("清除本檔斷點資料（強制重跑）", use_container_width=True):
                try:
                    for p in paths.values():
                        if isinstance(p, str) and os.path.exists(p):
                            os.remove(p)
                except Exception:
                    pass
                st.success("已清除本檔斷點資料")
                st.rerun()

        if auto_clean:
            try:
                media_dir = os.path.join(WORK_DIR, "media")
                if os.path.isdir(media_dir):
                    shutil.rmtree(media_dir, ignore_errors=True)
                st.toast("已清除暫存媒體檔案", icon="🧹")
            except Exception:
                pass

    except Exception as e:
        set_status("err", f"執行過程中發生錯誤：{e}")
        with st.expander("查看詳細錯誤資訊"):
            st.code(traceback.format_exc())


# =========================================================
#                      Header
# =========================================================
st.markdown(
    f"""
<div class="auro-header">
  <img src="{LOGO_URL}" alt="AUROTEK" />
  <div class="auro-subtitle">簡報案例自動化發布平台</div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="callout blue">
功能說明：上傳簡報 → 線上拆分 → 影片雲端化 → 內嵌優化 → 雲端發布 → 寫入和椿資料庫
</div>
""",
    unsafe_allow_html=True,
)

# =========================================================
#                  Session init / Bot init
# =========================================================
ensure_workspace()

if "split_jobs" not in st.session_state:
    st.session_state.split_jobs = []

if "ppt_meta" not in st.session_state:
    st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}

if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = None

if "bot" not in st.session_state:
    try:
        st.session_state.bot = PPTAutomationBot()
    except Exception as e:
        st.session_state.bot = None
        st.markdown(f"<div class='callout err'>Bot 初始化失敗：{e}</div>", unsafe_allow_html=True)

bot = st.session_state.bot
category_options = get_category_options(bot) if bot else ["清潔", "配送", "購物", "AURO"]

# =========================================================
#                       Step 1
# =========================================================
with st.container():
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("步驟一：選擇檔案來源")

    input_method = st.radio("上傳方式", ["本地檔案", "線上檔案"], horizontal=True)

    source_path = os.path.join(WORK_DIR, "source.pptx")
    file_name_for_logic = None

    if input_method == "本地檔案":
        uploaded_file = st.file_uploader("PPTX", type=["pptx"], label_visibility="collapsed")
        if uploaded_file:
            file_name_for_logic = uploaded_file.name
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
                    ok, err = download_file_from_url(url_input, source_path)
                    if ok:
                        file_name_for_logic = fake_name
                        st.toast("下載完成", icon="✅")
                    else:
                        st.markdown(f"<div class='callout err'>下載失敗：{err}</div>", unsafe_allow_html=True)

    # Parse and prepare preview + job restore
    if file_name_for_logic and os.path.exists(source_path):
        if st.session_state.current_file_name != file_name_for_logic:
            st.session_state.current_file_name = file_name_for_logic

            saved_jobs = load_history(file_name_for_logic)
            st.session_state.split_jobs = saved_jobs if saved_jobs else []

            ph = st.empty()
            ph.progress(0, text="解析簡報…")
            try:
                prs = Presentation(source_path)
                total_slides = len(prs.slides)

                preview_data = []
                for i, slide in enumerate(prs.slides):
                    txt = "無標題"
                    try:
                        if slide.shapes.title and getattr(slide.shapes.title, "text", "").strip():
                            txt = slide.shapes.title.text.strip()
                        else:
                            for s in slide.shapes:
                                if hasattr(s, "text"):
                                    t = (s.text or "").strip()
                                    if t:
                                        txt = t[:24] + ("…" if len(t) > 24 else "")
                                        break
                    except Exception:
                        pass
                    preview_data.append({"頁碼": i + 1, "內容摘要": txt})

                st.session_state.ppt_meta["total_slides"] = total_slides
                st.session_state.ppt_meta["preview_data"] = preview_data
                ph.progress(100, text="完成")

                st.markdown(
                    f"<div class='callout blue'>已讀取：{file_name_for_logic}（共 {total_slides} 頁）</div>",
                    unsafe_allow_html=True,
                )

                if not st.session_state.split_jobs:
                    # ✅ 修正：至少 1 頁也能正常給預設任務
                    st.session_state.split_jobs = [new_job_template(total_slides, category_options)]
                    save_history(file_name_for_logic, st.session_state.split_jobs)

            except Exception as e:
                st.markdown(f"<div class='callout err'>檔案處理失敗：{e}</div>", unsafe_allow_html=True)
                st.stop()

    st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
#                       Step 2
# =========================================================
if st.session_state.current_file_name:
    total_slides = st.session_state.ppt_meta.get("total_slides", 0)
    preview_data = st.session_state.ppt_meta.get("preview_data", [])

    with st.expander("頁碼與標題對照表", expanded=False):
        st.dataframe(preview_data, use_container_width=True, height=260, hide_index=True)

    with st.container():
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        left, right = st.columns([3, 1])
        left.subheader("步驟二：設定拆分任務")

        if right.button("新增任務", type="primary", use_container_width=True):
            add_split_job(total_slides, category_options)
            save_history(st.session_state.current_file_name, st.session_state.split_jobs)
            st.rerun()

        if not st.session_state.split_jobs:
            st.markdown("<div class='callout gray'>尚未建立任務，請先新增任務並設定頁數範圍。</div>", unsafe_allow_html=True)

        # ✅ 修正：total_slides 可能暫時為 0，先 clamp 到至少 1 避免 number_input 當機
        total_slides = max(1, int(total_slides or 0))

        for i, job in enumerate(st.session_state.split_jobs):
            # ✅ 修正：每次 render 前，把 start/end clamp 到合法範圍（避免 value 超出 min/max）
            try:
                job["start"] = max(1, min(int(job.get("start", 1)), total_slides))
            except Exception:
                job["start"] = 1
            try:
                job["end"] = max(1, min(int(job.get("end", total_slides)), total_slides))
            except Exception:
                job["end"] = total_slides
            if int(job["start"]) > int(job["end"]):
                job["start"] = int(job["end"])

            with st.container(border=True):
                st.markdown(f"**任務 {i+1}**")

                c1, c2, c3 = st.columns([3, 1.4, 1.4])
                job["filename"] = c1.text_input(
                    "檔名",
                    value=job.get("filename", ""),
                    key=f"fn_{job['id']}",
                    placeholder="例如：清潔案例A",
                )
                job["start"] = c2.number_input(
                    "起始頁",
                    min_value=1,
                    max_value=total_slides,
                    value=int(job.get("start", 1)),
                    key=f"st_{job['id']}",
                )
                job["end"] = c3.number_input(
                    "結束頁",
                    min_value=1,
                    max_value=total_slides,
                    value=int(job.get("end", total_slides)),
                    key=f"ed_{job['id']}",
                )

                m1, m2, m3, m4 = st.columns(4)

                current_cat = job.get("category", category_options[0] if category_options else "清潔")
                if current_cat not in category_options:
                    opts = [current_cat] + [x for x in category_options if x != current_cat]
                else:
                    opts = category_options

                job["category"] = m1.selectbox(
                    "類型",
                    options=opts,
                    index=0 if opts else 0,
                    key=f"cat_{job['id']}",
                )
                job["subcategory"] = m2.text_input("子分類", value=job.get("subcategory", ""), key=f"sub_{job['id']}")
                job["client"] = m3.text_input("客戶", value=job.get("client", ""), key=f"cli_{job['id']}")
                job["keywords"] = m4.text_input("關鍵字", value=job.get("keywords", ""), key=f"kw_{job['id']}")

                col_del, col_hint = st.columns([1, 5])
                with col_del:
                    if st.button("刪除任務", key=f"del_{job['id']}", type="secondary", use_container_width=True):
                        remove_split_job(i)
                        save_history(st.session_state.current_file_name, st.session_state.split_jobs)
                        st.rerun()
                with col_hint:
                    st.caption("提示：任務會從上方開始新增；若誤按可立即刪除。")

        save_history(st.session_state.current_file_name, st.session_state.split_jobs)
        st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
#                       Step 3
# =========================================================
if st.session_state.current_file_name:
    with st.container():
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("步驟三：開始執行")

        auto_clean = st.checkbox("任務完成後自動清除暫存媒體檔", value=True)

        src_fp = file_fingerprint(os.path.join(WORK_DIR, "source.pptx")) if os.path.exists(os.path.join(WORK_DIR, "source.pptx")) else ""
        can_resume, _info = detect_resume(st.session_state.current_file_name, src_fp) if src_fp else (False, {})
        if can_resume:
            st.markdown("<div class='callout blue'>偵測到斷點資料：可在網路中斷後接續執行（同檔名 + 同檔案指紋）。</div>", unsafe_allow_html=True)

        if st.button("執行自動化排程", type="primary", use_container_width=True):
            if not st.session_state.split_jobs:
                st.markdown("<div class='callout err'>請至少設定一個拆分任務後再執行。</div>", unsafe_allow_html=True)
            else:
                total_slides = max(1, int(st.session_state.ppt_meta.get("total_slides", 0) or 0))
                errs = validate_jobs(st.session_state.split_jobs, total_slides)
                if errs:
                    for e in errs:
                        st.error(e)
                    st.markdown("<div class='callout err'>請修正上述錯誤後再執行。</div>", unsafe_allow_html=True)
                else:
                    if not bot:
                        st.markdown("<div class='callout err'>機器人未初始化（憑證錯誤），請檢查 Secrets。</div>", unsafe_allow_html=True)
                        st.stop()

                    scroll_to_bottom()

                    execute_automation_logic(
                        bot=bot,
                        source_path=os.path.join(WORK_DIR, "source.pptx"),
                        file_name=st.session_state.current_file_name,
                        file_prefix=os.path.splitext(st.session_state.current_file_name)[0],
                        jobs=st.session_state.split_jobs,
                        auto_clean=auto_clean,
                    )

        st.markdown("</div>", unsafe_allow_html=True)
