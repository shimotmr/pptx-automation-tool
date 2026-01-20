# Version: v1.3 (Final Polish)
# Fixes:
# 1. Logic: Step 3 only appears AFTER tasks are added.
# 2. UI: All Green alerts changed to Blue (Info).
# 3. UI: Removed ALL Emojis from headers.
# 4. Layout: Trash button restored to top-right inline position.
# 5. Layout: Footer buttons aligned perfectly (removed extra HTML wrappers).
# 6. Spacing: Removed manual margins to ensure consistent spacing between steps.

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
    page_icon="📄",
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
header[data-testid="stHeader"] { display: none; }
.stApp > header { display: none; }

/* 2. 調整頂部與底部間距 */
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 6rem !important;
}

/* 3. 上傳按鈕樣式 (使用驗證過的 :not 排除法，確保無雙重按鈕) */
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

/* 鎖定主要按鈕 */
section[data-testid="stFileUploaderDropzone"] button {
    border: 1px solid #d0d7de;
    background-color: #ffffff;
    color: transparent !important; /* 隱藏英文 */
    position: relative;
    padding: 0.25rem 0.75rem;
    border-radius: 4px;
    min-height: 38px;
    width: auto;
    margin-top: 10px;
}

/* 疊加中文文字 */
section[data-testid="stFileUploaderDropzone"] button::after {
    content: "瀏覽檔案";
    position: absolute;
    color: #31333F;
    left: 50%; top: 50%;
    transform: translate(-50%, -50%);
    white-space: nowrap;
    font-weight: 500;
    font-size: 14px;
}

/* 排除刪除按鈕 (X) */
[data-testid="stFileUploaderDeleteBtn"] {
    border: none !important;
    background: transparent !important;
    margin-top: 0 !important;
    min-height: auto !important;
    color: inherit !important;
}
[data-testid="stFileUploaderDeleteBtn"]::after { content: none !important; }

/* 4. 統一字體與標題樣式 (無 Emoji) */
h3 { font-size: 1.2rem !important; font-weight: 700 !important; color: #31333F; margin-bottom: 0.5rem;}
h4 { font-size: 1.1rem !important; font-weight: 600 !important; color: #555; }
.stProgress > div > div > div > div { color: white; font-weight: 500; }

/* 5. 統一提示詞顏色 (強制藍色風格，覆蓋綠色) */
div[data-testid="stAlert"][data-style="success"],
div[data-testid="stAlert"][data-style="info"] {
    background-color: #F0F2F6 !important;
    color: #31333F !important;
    border: 1px solid #d0d7de !important;
}
div[data-testid="stAlert"] svg { color: #004280 !important; }
[data-testid="stAlert"] p { font-size: 0.9rem !important; line-height: 1.4 !important; }

/* 6. 垃圾桶按鈕樣式 (微調以適應欄位) */
div[data-testid="column"] button {
   border: 1px solid #eee !important;
   background: white !important;
   color: #555 !important;
   font-size: 0.85rem !important;
   white-space: nowrap !important;
   min-width: 40px !important; 
   padding: 4px 8px !important;
}
div[data-testid="column"] button:hover {
   color: #cc0000 !important;
   border-color: #cc0000 !important;
   background: #fff5f5 !important;
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
            print(f"Cleanup
