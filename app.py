import os
import uuid
import json
import shutil
import traceback
import requests

import streamlit as st
import streamlit.components.v1 as components
from pptx import Presentation

from ppt_processor import PPTAutomationBot


# =========================
# Config
# =========================
st.set_page_config(
    page_title="Aurotek數位資料庫 簡報案例自動化發布平台",
    page_icon="🤖",
    layout="wide",
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"
SOURCE_FILENAME = "source.pptx"


# =========================
# CSS
# =========================
st.markdown(
    """
<style>
/* 隱藏 Streamlit 預設 header */
header[data-testid="stHeader"] { display:none; }
.stApp > header { display:none; }

/* 整體間距更緊湊 */
.block-container{
  padding-top: 0.9rem !important;
  padding-bottom: 1.2rem !important;
}

/* ========= File Uploader（瘦身、中文化、修正直排重複） ========= */

/* 隱藏原本兩行說明 */
[data-testid="stFileUploaderDropzoneInstructions"] > div:first-child { display:none !important; }
[data-testid="stFileUploaderDropzoneInstructions"] > div:nth-child(2) { display:none !important; }

/* 用更短的中文提示 */
[data-testid="stFileUploaderDropzoneInstructions"]::before{
  content:"拖放或點擊上傳";
  display:block;
  font-size:0.95rem;
  font-weight:700;
  line-height:1.15;
  margin:0;
}
[data-testid="stFileUploaderDropzoneInstructions"]::after{
  content:"單檔 5GB · PPTX";
  display:block;
  font-size:0.75rem;
  color:#8a8a8a;
  margin-top:2px;
  line-height:1.15;
}

/* 讓 dropzone 更矮 */
section[data-testid="stFileUploaderDropzone"]{
  padding:0.55rem 0.9rem !important;
}

/* ====== 這段是「縱排重複瀏覽檔案」的根治 ======
   以前用 color: transparent 可能無法蓋掉內部 span，窄寬會變直排。
   改用 font-size:0 徹底讓原文字消失，再用 ::after 放中文。 */
div[data-testid="stFileUploader"] button{
  position:relative !important;
  font-size:0 !important;            /* ✅ 原本文字徹底消失（避免直排） */
  line-height:0 !important;
  white-space:nowrap !important;
  writing-mode: horizontal-tb !important;
}
div[data-testid="stFileUploader"] button::after{
  content:"瀏覽檔案";
  font-size:0.95rem;
  line-height:1;
  color:#31333F;
  font-weight:600;
  position:absolute;
  left:50%; top:50%;
  transform:translate(-50%, -50%);
  white-space:nowrap;
  writing-mode: horizontal-tb;
}

/* st.info 文字稍微小一點 */
[data-testid="stAlert"] p{
  font-size:0.85rem !important;
  line-height:1.35 !important;
}

/* ===== Results UI ===== */
.auro-result-wrap{
  border:1px solid rgba(49,51,63,.15);
  border-radius:14px;
  padding:14px 14px 8px 14px;
  background:#fff;
}
.auro-result-title{
  display:flex;
  align-items:center;
  gap:10px;
  font-size:1.15rem;
  font-weight:800;
  margin:0 0 8px 0;
}
.auro-pill{
  display:inline-block;
  padding:2px 10px;
  border-radius:999px;
  font-size:0.78rem;
  color:#0b5;
  background:rgba(0,170,85,.10);
  border:1px solid rgba(0,170,85,.25);
}
.auro-card{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:10px;
  padding:12px 12px;
  margin:10px 0;
  border:1px solid rgba(49,51,63,.12);
  border-radius:12px;
  background:rgba(248,249,251,.7);
}
.auro-card .name{
  font-weight:700;
  color:#222;
  overflow:hidden;
  text-overflow:ellipsis;
  white-space:nowrap;
  max-width:70vw;
}
.auro-card a.btn{
  text-decoration:none !important;
  padding:8px 12px;
  border-radius:10px;
  background:#0B4F8A;
  color:white !important;
  font-weight:700;
  white-space:nowrap;
}
.auro-card a.btn:hover{
  filter:brightness(1.05);
}

/* 手機再緊一點 */
@media (max-width:768px){
  .block-container{ padding-top:0.65rem !important; }
  .auro-card .name{ max-width:55vw; }
}
</style>
""",
    unsafe_allow_html=True,
)


# =========================
# Utilities
# =========================
def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def clean_workspace() -> None:
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR, ignore_errors=True)
    ensure_dir(WORK_DIR)


def write_source_to_workspace(file_bytes: bytes) -> str:
    """回傳 source.pptx 的實際路徑"""
    clean_workspace()
    source_path = os.path.join(WORK_DIR, SOURCE_FILENAME)
    with open(source_path, "wb") as f:
        f.write(file_bytes)
    return source_path


def load_history(filename: str):
    if not os.path.exists(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get(filename, [])
    except Exception:
        return []


def save_history(filename: str, jobs):
    try:
        data = {}
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}
        data[filename] = jobs
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"History save failed: {e}")


def add_split_job(total_pages: int):
    st.session_state.split_jobs.insert(
        0,
        {
            "id": str(uuid.uuid4())[:8],
            "filename": "",
            "start": 1,
            "end": total_pages,
            "category": "清潔",
            "subcategory": "",
            "client": "",
            "keywords": "",
        },
    )


def remove_split_job(index: int):
    st.session_state.split_jobs.pop(index)


def validate_jobs(jobs, total_slides: int):
    errors = []
    for i, job in enumerate(jobs):
        task_label = f"任務 {i+1} (檔名: {job['filename'] or '未命名'})"
        if not job["filename"].strip():
            errors.append(f"❌ {task_label}: 檔案名稱不能為空。")
        if job["start"] > job["end"]:
            errors.append(f"❌ {task_label}: 起始頁 ({job['start']}) 不能大於 結束頁 ({job['end']})。")
        if job["end"] > total_slides:
            errors.append(f"❌ {task_label}: 結束頁 ({job['end']}) 超出了簡報總頁數 ({total_slides})。")

    sorted_jobs = sorted(jobs, key=lambda x: x["start"])
    for i in range(len(sorted_jobs) - 1):
        a, b = sorted_jobs[i], sorted_jobs[i + 1]
        if a["end"] >= b["start"]:
            errors.append(
                "⚠️ 發現頁數重疊！\n"
                f"   - {a['filename']} (範圍 {a['start']}-{a['end']})\n"
                f"   - {b['filename']} (範圍 {b['start']}-{b['end']})\n"
                f"   請確認是否重複包含了第 {b['start']} 到 {a['end']} 頁。"
            )
    return errors


def download_bytes_from_url(url: str):
    r = requests.get(url, stream=True, timeout=60)
    r.raise_for_status()
    return r.content


@st.cache_resource(show_spinner=False)
def get_bot():
    return PPTAutomationBot()


@st.cache_data(show_spinner=False)
def parse_ppt_preview(ppt_bytes: bytes):
    """回傳 (total_slides, preview_data)"""
    ensure_dir(WORK_DIR)
    tmp_path = os.path.join(WORK_DIR, "__preview__.pptx")
    with open(tmp_path, "wb") as f:
        f.write(ppt_bytes)

    prs = Presentation(tmp_path)
    total = len(prs.slides)
    preview_data = []
    for i, slide in enumerate(prs.slides):
        txt = "無標題"
        try:
            if slide.shapes.title and slide.shapes.title.text:
                txt = slide.shapes.title.text
        except Exception:
            pass
        if txt == "無標題":
            for s in slide.shapes:
                if hasattr(s, "text") and isinstance(s.text, str) and s.text.strip():
                    txt = s.text.strip()[:20] + "..."
                    break
        preview_data.append({"頁碼": i + 1, "內容摘要": txt})
    return total, preview_data


def render_header():
    # 高度改得比較緊凑
    components.html(
        f"""
        <div style="
            width:100%;
            display:flex;
            flex-direction:column;
            align-items:center;
            justify-content:center;
            margin:2px 0 0 0;
            line-height:1.05;
        ">
            <img src="{LOGO_URL}" style="width:300px;height:auto;display:block;margin:0;" />
            <div style="
                margin-top:4px;
                width:300px;
                text-align:center;
                color:gray;
                font-size:1.0rem;
                font-weight:500;
                letter-spacing:2px;
            ">簡報案例自動化發布平台</div>
        </div>
        """,
        height=74,
    )


def render_results_ui(file_prefix: str, final_results: list):
    links = []
    for res in final_results:
        if "final_link" in res:
            links.append((f"[{file_prefix}]_{res.get('filename','')}", res["final_link"]))

    st.markdown('<div class="auro-result-wrap">', unsafe_allow_html=True)
    st.markdown(
        f'<div class="auro-result-title">✅ 產出結果連結 <span class="auro-pill">{len(links)} 筆</span></div>',
        unsafe_allow_html=True,
    )

    if not links:
        st.info("沒有產生任何結果連結，請檢查是否有任務被跳過。")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    for name, link in links:
        st.markdown(
            f"""
            <div class="auro-card">
              <div class="name">{name}</div>
              <a class="btn" href="{link}" target="_blank" rel="noopener">開啟 Google Slides</a>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("</div>", unsafe_allow_html=True)


# =========================
# Core automation
# =========================
def execute_automation_logic(bot, source_path, file_prefix, jobs, auto_clean):
    main_progress = st.progress(0, text="準備開始...")
    status_area = st.empty()
    detail_bar = st.empty()

    sorted_jobs = sorted(jobs, key=lambda x: x["start"])

    def update_detail(pct, text):
        detail_bar.progress(pct, text=text)

    def log(msg):
        print(f"[Log] {msg}")

    try:
        status_area.info("1️⃣ 步驟 1/5：提取 PPT 內影片並上傳至雲端...")
        main_progress.progress(5, text="Step 1: 影片雲端化")
        video_map = bot.extract_and_upload_videos(
            source_path,
            os.path.join(WORK_DIR, "media"),
            file_prefix=file_prefix,
            progress_callback=lambda fn, cur, tot: update_detail(
                cur / tot if tot else 0,
                f"Step 1：上傳 `{fn}` ({int((cur/tot)*100) if tot else 0}%)",
            ),
            log_callback=log,
        )
        detail_bar.empty()

        status_area.info("2️⃣ 步驟 2/5：將 PPT 內的影片替換為雲端連結圖片...")
        main_progress.progress(25, text="Step 2: 連結置換")
        mod_path = os.path.join(WORK_DIR, "modified.pptx")
        bot.replace_videos_with_images(
            source_path,
            mod_path,
            video_map,
            progress_callback=lambda cur, tot: update_detail(
                cur / tot if tot else 0,
                f"Step 2：處理投影片 {cur}/{tot}",
            ),
        )
        detail_bar.empty()

        status_area.info("3️⃣ 步驟 3/5：進行檔案壓縮與瘦身...")
        main_progress.progress(45, text="Step 3: 檔案瘦身")
        slim_path = os.path.join(WORK_DIR, "slim.pptx")
        bot.shrink_pptx(
            mod_path,
            slim_path,
            progress_callback=lambda cur, tot: update_detail(
                cur / tot if tot else 0,
                f"Step 3：處理內部檔案 {cur}/{tot}",
            ),
        )
        detail_bar.empty()

        status_area.info("4️⃣ 步驟 4/5：依設定拆分簡報並上傳至 Google Slides...")
        main_progress.progress(65, text="Step 4: 拆分發布")
        results = bot.split_and_upload(
            slim_path,
            sorted_jobs,
            file_prefix=file_prefix,
            progress_callback=lambda fn, cur, tot: update_detail(
                cur / tot if tot else 0,
                f"Step 4：上傳 `{fn}` ({int((cur/tot)*100) if tot else 0}%)",
            ),
            log_callback=log,
        )
        detail_bar.empty()

        oversized = [r for r in results if r.get("error_too_large")]
        if oversized:
            st.error("⛔️ 流程終止：偵測到拆分後的檔案過大（超過 Google 100MB 限制）。")
            for j in oversized:
                st.error(f"❌ 任務「{j['filename']}」仍有 {j['size_mb']:.2f} MB")
            st.warning("💡 建議：縮小該任務頁數範圍，拆成多個小任務後再跑。")
            return

        status_area.info("5️⃣ 步驟 5/5：優化線上簡報的影片播放器...")
        main_progress.progress(85, text="Step 5: 內嵌優化")
        final_results = bot.embed_videos_in_slides(
            results,
            progress_callback=lambda cur, tot: update_detail(
                cur / tot if tot else 0,
                f"Step 5：優化任務 {cur}/{tot}",
            ),
            log_callback=log,
        )
        detail_bar.empty()

        status_area.info("📝 最後步驟：將成果寫入 Google Sheets 資料庫...")
        main_progress.progress(95, text="Final: 寫入資料庫")
        bot.log_to_sheets(final_results, log_callback=log)

        main_progress.progress(100, text="🎉 任務全部完成！")
        status_area.success("🎉 所有自動化流程執行完畢！")

        if auto_clean:
            clean_workspace()
            st.toast("已自動清除暫存檔案。", icon="🧹")

        st.divider()
        render_results_ui(file_prefix, final_results)

    except Exception as e:
        st.error(f"執行過程中發生錯誤: {e}")
        with st.expander("查看詳細錯誤資訊"):
            st.code(traceback.format_exc())


# =========================
# State init
# =========================
ensure_dir(WORK_DIR)

if "split_jobs" not in st.session_state:
    st.session_state.split_jobs = []
if "ppt_meta" not in st.session_state:
    st.session_state.ppt_meta = {"total_slides": 0, "preview_data": []}
if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = None
if "source_bytes" not in st.session_state:
    st.session_state.source_bytes = None


# =========================
# UI
# =========================
render_header()

st.info("功能說明： 上傳PPT → 線上拆分 → 影片雲端化 → 內嵌優化 → 簡報雲端化 → 寫入和椿資料庫")

# Bot init（快取）
try:
    bot = get_bot()
    if not getattr(bot, "creds", None):
        st.warning("⚠️ 系統未檢測到有效憑證 (Secrets)。")
except Exception as e:
    st.error(f"Bot 初始化失敗: {e}")
    bot = None

# Step 1
with st.container(border=True):
    st.subheader("📂 步驟一：選擇檔案來源")

    input_method = st.radio("上傳方式", ["本地檔案", "線上檔案"], horizontal=True)

    file_name_for_logic = None

    if input_method == "本地檔案":
        uploaded_file = st.file_uploader("PPTX", type=["pptx"], label_visibility="collapsed")
        if uploaded_file:
            st.session_state.source_bytes = uploaded_file.getvalue()
            file_name_for_logic = uploaded_file.name

    else:
        url_input = st.text_input("PPTX 直接下載網址 (Direct URL)", placeholder="https://example.com/file.pptx")
        if url_input:
            if not url_input.lower().endswith(".pptx"):
                st.warning("⚠️ 網址結尾似乎不是 .pptx，請確認網址正確性。")

            fake_name = url_input.split("/")[-1].split("?")[0]
            if not fake_name.lower().endswith(".pptx"):
                fake_name += ".pptx"

            if st.button("📥 下載並載入", use_container_width=True):
                with st.spinner("正在下載檔案..."):
                    try:
                        st.session_state.source_bytes = download_bytes_from_url(url_input)
                        file_name_for_logic = fake_name
                        st.success("下載成功！")
                    except Exception as e:
                        st.error(f"下載失敗: {e}")

    # 解析 PPT（當檔名變更才重新讀）
    if file_name_for_logic and st.session_state.source_bytes:
        if st.session_state.current_file_name != file_name_for_logic:
            # 寫入 workspace（✅ 不會再發生 source.pptx 被 cleanup 刪掉造成 Package not found）
            source_path = write_source_to_workspace(st.session_state.source_bytes)

            # 載入歷史任務
            st.session_state.split_jobs = load_history(file_name_for_logic) or []

            # 解析頁面（cache）
            with st.spinner("解析檔案中..."):
                try:
                    total_slides, preview_data = parse_ppt_preview(st.session_state.source_bytes)
                    st.session_state.ppt_meta = {"total_slides": total_slides, "preview_data": preview_data}
                    st.session_state.current_file_name = file_name_for_logic
                    st.success(f"✅ 已讀取：{file_name_for_logic} (共 {total_slides} 頁)")
                except Exception as e:
                    st.error(f"檔案處理失敗: {e}")
                    st.session_state.current_file_name = None

# Step 2/3（僅當已載入）
if st.session_state.current_file_name:
    total_slides = st.session_state.ppt_meta["total_slides"]
    preview_data = st.session_state.ppt_meta["preview_data"]

    with st.expander("👁️ 點擊查看「頁碼與標題對照表」", expanded=False):
        st.dataframe(preview_data, use_container_width=True, height=250, hide_index=True)

    with st.container(border=True):
        head_l, head_r = st.columns([3, 1])
        head_l.subheader("📝 步驟二：設定拆分任務")
        if head_r.button("➕ 新增任務", type="primary", use_container_width=True):
            add_split_job(total_slides)

        if not st.session_state.split_jobs:
            st.info("☝️ 尚未建立任務，請點擊上方按鈕新增。")

        for i, job in enumerate(st.session_state.split_jobs):
            with st.container(border=True):
                st.markdown(f"**📄 任務 {i+1}**")
                c1, c2, c3 = st.columns([3, 1.3, 1.3])
                job["filename"] = c1.text_input("檔名", value=job["filename"], key=f"f_{job['id']}", placeholder="例如：清潔案例A")
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

        save_history(st.session_state.current_file_name, st.session_state.split_jobs)

    with st.container(border=True):
        st.subheader("🚀 開始執行")
        auto_clean = st.checkbox("任務完成後自動清除暫存檔", value=True)

        if st.button("執行自動化排程", type="primary", use_container_width=True):
            if not st.session_state.split_jobs:
                st.error("請至少設定一個拆分任務！")
            else:
                errs = validate_jobs(st.session_state.split_jobs, total_slides)
                if errs:
                    for e in errs:
                        st.error(e)
                    st.error("⛔️ 請修正錯誤後再執行。")
                else:
                    if not bot:
                        st.error("❌ 機器人未初始化（Secrets/憑證問題）。")
                    else:
                        source_path = os.path.join(WORK_DIR, SOURCE_FILENAME)
                        file_prefix = os.path.splitext(st.session_state.current_file_name)[0]
                        execute_automation_logic(bot, source_path, file_prefix, st.session_state.split_jobs, auto_clean)
