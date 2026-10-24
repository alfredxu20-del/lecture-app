"""
Streamlit 讲座信息汇总 Web 应用
================================
功能:
1. 上传讲座截图（PNG / JPG / JPEG / WEBP）
2. 调用大模型视觉 API 提取讲座信息（JSON）
3. 写入 SQLite 数据库（按 title+time_start 去重）
4. 下方表格展示所有讲座，支持筛选 / 搜索 / 导出 Excel

运行:
  streamlit run app.py
"""

import os
import re
import io
import json
import time
import base64
import sqlite3
import uuid
from datetime import datetime
from dotenv import load_dotenv
import libsql
from typing import Optional

import pandas as pd
import streamlit as st
from PIL import Image
import requests
from lxml import html as LH

# ============== 配置区（从 .env 读取，不提交） ==============
load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

# 优先级：st.secrets（Streamlit Cloud）> 环境变量（本地 .env）
def _safe_secret(key, default=""):
    """先试 st.secrets，失败回退 os.getenv。

    关键：Streamlit Cloud 的 Secrets 只能通过 st.secrets 读取，
    且云端并不存在本地 .streamlit/secrets.toml 文件，
    所以这里必须无条件先尝试 st.secrets，不能先判断本地文件是否存在。
    """
    try:
        return st.secrets[key]
    except Exception:
        return os.getenv(key, default)

API_KEY    = _safe_secret("DEEPSEEK_API_KEY")
API_BASE   = _safe_secret("DEEPSEEK_API_BASE", "https://api.deepseek.com/v1")
MODEL_NAME = _safe_secret("DEEPSEEK_MODEL", "deepseek-chat")

# 数据库：有 Turso 环境变量走云端，否则本地 SQLite
TURSO_URL    = _safe_secret("TURSO_DB_URL").strip()
TURSO_TOKEN  = _safe_secret("TURSO_AUTH_TOKEN").strip()
USE_TURSO    = bool(TURSO_URL and TURSO_TOKEN)
LOCAL_DB     = os.path.join(os.path.dirname(__file__), "lectures.db")
EXCEL_PATH   = os.path.join(os.path.dirname(__file__), "lectures.xlsx")
UPLOAD_DIR   = os.path.join(os.path.dirname(__file__), "uploads")

# 访客录入口令
GUEST_PASSCODE = _safe_secret("GUEST_PASSCODE").strip()

# ============== 页面初始化 ==============
st.set_page_config(
    page_title="讲座信息汇总 · 复旦哲学",
    layout="wide",
)

# ============== 自定义 CSS（清冷浅节风格） ==============
# 需要的素材文件（放在 D:\lecture-app\assets\ 下）：
#   owl.png   —— 猫头鹰线稿水印，建议尺寸 600x600，PNG 透明背景，单色（#8A9AAB 或 #4A5568），笔画细
#   fudan.png —— 复旦校徽，建议尺寸 400x400，PNG 透明背景，单色线稿
#   school.png —— 哲学学院院徽，建议尺寸 400x400，PNG 透明背景，单色线稿
# 图片处理：用任意图像工具把底色调成灰白/雾霾蓝，再去掉背景即可。
# 如果文件不存在，下面的 url() 引用会优雅降级（不显示水印，不报错）。
import base64
from pathlib import Path as _Path

_ASSETS_DIR = _Path(__file__).parent / "assets"

def _b64_or_none(name: str) -> str | None:
    p = _ASSETS_DIR / name
    if not p.exists():
        return None
    return base64.b64encode(p.read_bytes()).decode("ascii")

_owl_b64    = _b64_or_none("owl.png")
_fudan_b64  = _b64_or_none("fudan.png")
_school_b64 = _b64_or_none("school.png")

_bg_layers = []
if _owl_b64:
    _bg_layers.append(f"url(data:image/png;base64,{_owl_b64})")
if _fudan_b64:
    _bg_layers.append(f"url(data:image/png;base64,{_fudan_b64})")
if _school_b64:
    _bg_layers.append(f"url(data:image/png;base64,{_school_b64})")
_bg_image_css = ", ".join(_bg_layers) if _bg_layers else "none"

_corner_logo_b64 = _fudan_b64 or _school_b64 or _owl_b64
_corner_logo_css = (
    f"url(data:image/png;base64,{_corner_logo_b64})" if _corner_logo_b64 else "none"
)

st.markdown(
    f"""
<style>
/* 品牌：衬线字体。其余视觉 token 由 .streamlit/config.toml 控制 */
html, body, [data-testid="stApp"] {{
  font-family: "Source Han Serif SC", "Noto Serif SC", "Songti SC",
               "SimSun", "Source Han Sans SC", "PingFang SC", sans-serif;
}}

/* 品牌：背景水印（猫头鹰 + 校徽，低透明度铺底）—— config.toml 无法实现 ::before 背景图 */
[data-testid="stApp"]::before {{
  content: "";
  position: fixed;
  inset: 0;
  pointer-events: none;
  z-index: 0;
  background-image: {_bg_image_css};
  background-repeat: no-repeat, no-repeat, no-repeat;
  background-position: center 120%, right -120px top 40px, left -100px bottom 60px;
  background-size: 720px auto, 360px auto, 320px auto;
  opacity: 0.03;
}}

/* 让主容器浮在水印之上 */
[data-testid="stAppViewContainer"] > .main {{
  position: relative;
  z-index: 1;
}}

/* 品牌：顶角校徽 —— config.toml 无法实现 ::after 背景图 */
[data-testid="stAppViewContainer"]::after {{
  content: "";
  position: fixed;
  top: 18px;
  right: 22px;
  width: 56px;
  height: 56px;
  background-image: {_corner_logo_css};
  background-repeat: no-repeat;
  background-size: contain;
  opacity: 0.10;
  pointer-events: none;
  z-index: 2;
}}

/* 搜索框：纯白背景 + 1px 浅灰边框，直角，加宽加高，去阴影 */
[data-testid="stTextInput"] input {{
  background: #FFFFFF !important;
  border: 1px solid #D9DCE1 !important;
  border-radius: 0 !important;
  box-shadow: none !important;
  padding-top: 0.6rem !important;
  padding-bottom: 0.6rem !important;
  font-size: 1rem !important;
}}
[data-testid="stTextInput"] input:focus {{
  border-color: #4A5568 !important;
  box-shadow: none !important;
}}

/* 极简列表视图：1px 浅灰水平分割线，行内左右布局 */
.list-view {{
  margin: 0;
}}
/* 固定高度滚动容器：列表项过多时内部下拉，不撑长页面 */
.list-scroll {{
  max-height: 420px;
  overflow-y: auto;
  padding-right: 4px;
}}
.list-scroll::-webkit-scrollbar {{
  width: 6px;
}}
.list-scroll::-webkit-scrollbar-thumb {{
  background: #D9DCE1;
  border-radius: 3px;
}}
.list-scroll::-webkit-scrollbar-thumb:hover {{
  background: #B5BAC1;
}}
.list-item {{
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  padding: 12px 2px;
  border-bottom: 1px solid #EDEDEF;
  flex-wrap: wrap;
}}
.list-item:last-child {{
  border-bottom: none;
}}
.list-main {{
  flex: 1 1 60%;
  min-width: 0;
}}
.list-title {{
  font-size: 1rem;
  font-weight: 600;
  color: #2D3748;
  line-height: 1.45;
  margin: 0 0 4px 0;
  word-break: break-word;
}}
.list-meta {{
  font-size: 0.85rem;
  color: #8A8F98;
  line-height: 1.4;
  margin: 0;
  word-break: break-word;
}}
.list-src {{
  color: #8A8F98;
  text-decoration: none;
  border-bottom: 1px solid #C9CDD3;
  padding-bottom: 1px;
}}
.list-src:hover {{
  color: #2D3748;
  border-bottom-color: #2D3748;
}}
.list-side {{
  flex: 0 1 auto;
  max-width: 40%;
  text-align: right;
  font-size: 0.85rem;
  color: #8A8F98;
  line-height: 1.4;
  word-break: break-word;
  white-space: normal;
  overflow-wrap: anywhere;
}}
/* 窄屏：右侧信息换行到标题下方，左对齐 */
@media (max-width: 640px) {{
  .list-side {{
    flex-basis: 100%;
    max-width: 100%;
    text-align: left;
    color: #8A8F98;
  }}
}}
.list-empty {{
  text-align: center;
  color: #B5BAC1;
  padding: 28px 0;
  font-size: 0.9rem;
}}
</style>
""",
    unsafe_allow_html=True,
)


st.session_state.setdefault("auth_ok", False)

_head_l, _head_r = st.columns([4, 1])
with _head_l:
    st.markdown("## 讲座 · 人文活动 · 学习场所 汇总")
    st.caption("复旦大学哲学学院 ｜ 自动识别分类 · 汇总 · 检索")
with _head_r:
    st.write("")
    if st.button("导入新信息", type="primary", use_container_width=True, key="btn_import"):
        if st.session_state["auth_ok"]:
            st.session_state["_req_import"] = True
        else:
            st.session_state["_req_pass"] = True

# ============== 数据库工具
# ============== 数据库工具 ==============
SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS lectures (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    category        TEXT    DEFAULT 'lecture',
    title           TEXT    NOT NULL,
    speaker         TEXT,
    affiliation     TEXT,
    time_text       TEXT,
    time_start      TEXT,
    time_end        TEXT,
    location        TEXT,
    online_url      TEXT,
    host            TEXT,
    registration    TEXT,
    deadline        TEXT,
    source_account  TEXT,
    source_url      TEXT,
    source_type     TEXT,
    summary         TEXT,
    theme           TEXT,
    participants    TEXT,
    venue           TEXT,
    hours           TEXT,
    open_period     TEXT,
    contact         TEXT,
    confidence      REAL,
    created_at      TEXT    DEFAULT (datetime('now', 'localtime')),
    UNIQUE(title, time_start)
);
"""


def get_conn():
    """获取数据库连接：Turso 优先，否则本地 SQLite。返回统一接口 (conn, cursor_factory)。"""
    if USE_TURSO:
        # libsql 远程连接；多线程安全(Streamlit worker 共享)
        conn = libsql.connect(TURSO_URL, auth_token=TURSO_TOKEN)
        # 返回代理：cursor 默认按 tuple 返回（libsql 原生行为）
        return conn
    conn = sqlite3.connect(LOCAL_DB, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    """初始化数据库表，兼容老表新增列。"""
    conn = get_conn()
    try:
        conn.execute(SCHEMA_SQL)
        # 兼容：老表没有这些列则补上
        cur = conn.execute("PRAGMA table_info(lectures)")
        cols = [r[1] if not isinstance(r, dict) else r["name"] for r in cur.fetchall()]
        for col, ddl in [
            ("category",     "TEXT DEFAULT 'lecture'"),
            ("theme",        "TEXT"),
            ("participants", "TEXT"),
            ("venue",        "TEXT"),
            ("hours",        "TEXT"),
            ("open_period",  "TEXT"),
            ("contact",      "TEXT"),
            ("source_url",   "TEXT"),
            ("source_type",  "TEXT"),
            ("source_image", "TEXT"),
        ]:
            if col not in cols:
                conn.execute(f"ALTER TABLE lectures ADD COLUMN {col} {ddl}")
        conn.commit()
    finally:
        conn.close()


@st.cache_data(show_spinner=False)
def load_all_lectures() -> pd.DataFrame:
    """加载讲座到 DataFrame（缓存，刷新或新插入后失效）。"""
    conn = get_conn()
    try:
        df = pd.read_sql_query(
            "SELECT * FROM lectures ORDER BY time_start ASC", conn
        )
    finally:
        conn.close()
    # 防御性补列：旧缓存 DataFrame 可能缺少 category 等新增字段
    if "category" not in df.columns:
        df["category"] = "lecture"
    # 剥掉时区：避免表格里出现 "+08:00"，且不能误把上海时间转成 UTC
    # 正确做法：tz_convert("Asia/Shanghai") 保持字面时间，再 tz_localize(None) 丢时区标签
    # 注意：库里历史时间字符串格式曾混用 "YYYY-MM-DD HH:MM:SS" 与 "YYYY-MM-DDTHH:MM:SS"，
    # 向量化 to_datetime 对混合格式可能解析失败产生 NaT，故用 format="mixed" 逐值兜底。
    for _col in ("time_start", "time_end"):
        if _col in df.columns:
            df[_col] = pd.to_datetime(df[_col], errors="coerce", format="mixed")
            if df[_col].dt.tz is not None:
                df[_col] = df[_col].dt.tz_convert("Asia/Shanghai").dt.tz_localize(None)
    return df


def _normalize_key(s) -> str:
    """规范化字符串用于去重比对：去空白、全角转半角、大小写统一、去常见标点。"""
    if s is None:
        return ""
    s = str(s)
    # 全角转半角（含全角标点）
    out = []
    for ch in s:
        code = ord(ch)
        if code == 0x3000:
            out.append(" ")
        elif 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        else:
            out.append(ch)
    s = "".join(out)
    # 去所有空白与标点符号（保留字母数字与中文）
    s = re.sub(r"[\s\W_]+", "", s, flags=re.UNICODE)
    return s.lower()


def is_duplicate(record: dict) -> bool:
    """模糊去重：规范化 title + 同一 time_start 视为重复（在 DB UNIQUE 约束之外兜底）。"""
    title = record.get("title")
    ts = record.get("time_start")
    if not title or not ts:
        return False
    key = _normalize_key(title)
    if not key:
        return False
    conn = get_conn()
    try:
        cur = conn.execute(
            "SELECT title FROM lectures WHERE time_start = ?", (str(ts),)
        )
        rows = cur.fetchall()
    finally:
        conn.close()
    for (t,) in rows:
        if _normalize_key(t) == key:
            return True
    return False


def cleanup_expired() -> int:
    """物理删除已结束的讲座/人文活动（time_end 或 time_start 早于当前时间）。
    学习场所（study_venue）为长期开放信息，不清理。返回删除条数。"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_conn()
    try:
        cur = conn.execute(
            """
            DELETE FROM lectures
            WHERE category IN ('lecture', 'humanities_activity')
              AND COALESCE(NULLIF(time_end, ''), NULLIF(time_start, ''), '9999-12-31 00:00:00') < ?
            """,
            (now,),
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def insert_lecture(record: dict) -> tuple[bool, str]:
    """插入一条记录。返回 (是否新插入, 提示信息)。去重：title+time_start。"""
    sql = """
    INSERT OR IGNORE INTO lectures (
        category, title, speaker, affiliation,
        time_text, time_start, time_end,
        location, online_url, host, registration, deadline,
        source_account, source_url, source_type, source_image, summary,
        theme, participants,
        venue, hours, open_period, contact
    ) VALUES (
        :category, :title, :speaker, :affiliation,
        :time_text, :time_start, :time_end,
        :location, :online_url, :host, :registration, :deadline,
        :source_account, :source_url, :source_type, :source_image, :summary,
        :theme, :participants,
        :venue, :hours, :open_period, :contact
    )
    """
    # 补齐可能缺失的键
    for f in FIELDS:
        record.setdefault(f, None)
    # 规范化 category：未知 → lecture
    if record.get("category") not in CATEGORY_LIST:
        record["category"] = "lecture"
    # 模糊去重（在 DB UNIQUE 约束之外兜底：标题标点/空格/大小写差异）
    if is_duplicate(record):
        return False, "检测到疑似重复信息（标题与时间相同），已自动跳过。"
    conn = get_conn()
    try:
        # 顺序必须与 INSERT 列名一致；libsql 不支持 :name 命名参数，所以这里直接按列顺序传 list
        cols_order = [
            "category", "title", "speaker", "affiliation",
            "time_text", "time_start", "time_end",
            "location", "online_url", "host", "registration", "deadline",
            "source_account", "source_url", "source_type", "source_image", "summary",
            "theme", "participants",
            "venue", "hours", "open_period", "contact",
        ]
        params = [record.get(c) for c in cols_order]
        cur = conn.execute(sql, params)
        conn.commit()
        if cur.rowcount == 0:
            return False, "检测到疑似重复讲座，已自动跳过。"
        # 写入成功后自动清理已过期信息（物理删除）
        removed = cleanup_expired()
        tail = f"（本次已自动清理 {removed} 条已结束的信息）" if removed else ""
        return True, "已成功写入数据库。" + tail
    finally:
        conn.close()


# ============== 大模型视觉 API ==============
EXTRACTION_PROMPT = """
你是高校/学术机构的“信息识别助手”。请阅读这张海报或公众号文章截图，首先判断它属于以下哪一类，再严格按 JSON 输出：

【1】lecture            = 讲座 / 论坛 / 学术报告 / 研讨会 / 报告文学性质有主讲人的学术类活动
【2】humanities_activity = 人文活动 = 征稿 / 比赛 / 音乐节 / 艺术节 / 研习班 / 读书会 / 沙龙 / 院庆 / 校庆 / 志庆日 / 周年纪念 等系列活动（不是某位专家的单一学术讲座）
【3】study_venue        = 学习场所开放信息 = 阅览室 / 自习室 / 图书馆 / 实验室 开放时间、场地预约、调整公告

只输出一段 JSON，不要任何解释、不要 markdown 代码块、不要额外文字。JSON Schema：

{
  "category": "lecture" | "humanities_activity" | "study_venue",
  "title": "信息标题",
  "speaker":       "主讲人/主讲人列表（如是讲座）",
  "affiliation":   "主讲人单位（讲座类别）",
  "time_text":     "原文中的时间描述",
  "time_start":    "ISO8601 开始时间",
  "time_end":      "ISO8601 结束时间",
  "location":      "线下具体地点（讲座 / 活动适用）",
  "online_url":    "线上链接",
  "host":          "主办方",
  "registration":  "报名方式 / 报名链接 / 报名邮箱",
  "deadline":      "报名 / 投稿 / 申请 截止日期",
  "source_account":"公众号名",
  "summary":       "一句话摘要",
  "theme":         "人文活动的专题/主题（如：哲学与文学、音乐节主题、征稿主题）",
  "participants":  "人文活动适用：参与对象 / 参赛要求 / 参赛人群",
  "venue":         "学习场所适用：场所名称（如：理科图书馆三楼阅览室）",
  "hours":         "学习场所适用：开放时间描述（如：周一至周五 8:00-22:00）",
  "open_period":   "学习场所适用：有效起止日期（如：2026-10-01 ~ 2027-01-15）",
  "contact":       "学习场所适用：联系人 / 咨询电话 / 联系微信"
}

判断示例（供你参考，不是规则）：
- 主讲人明确、有讲座标题：“林彦明 | 当代哲学话语的重构” → lecture
- 征稿启事 / Call for Papers、主题创作大赛、青年音乐节、夏令营、冬令营、读书会、研习班 → humanities_activity
- 院庆 / 校庆 / 志庆日 / 周年纪念等系列活动（如“XX学院成立七十周年志庆日”），即使标题含“学术报告会 / 报告会”字样 → humanities_activity
- 图书馆阅览室调整开放时间、场馆闭馆通知、实验室开放预约 → study_venue

重要规则：
1. 三选一，判断不准时优先 lecture；但标题或正文含“院庆 / 校庆 / 志庆 / 周年纪念 / 志庆日”等纪念性质字样时，优先 humanities_activity。
2. 找不到的字段填 null，绝不能编造。
3. 时间统一 ISO8601（如 2026-10-15T14:00:00+08:00）；如果只有日期，time_start 填“YYYY-MM-DDT00:00:00”，time_end 同理。
4. 学习场所类别的时间通常不是一次“活动开始时间”，open_period/hours 填该场所的开放时段。
5. 只输出 JSON，不要任何额外文字。
"""


def encode_image_to_data_url(img: Image.Image) -> str:
    """把 PIL Image 编码成 base64 data URL。"""
    buf = io.BytesIO()
    fmt = img.format if img.format in ("PNG", "JPEG", "WEBP") else "PNG"
    if fmt == "PNG":
        img.save(buf, format="PNG")
        mime = "image/png"
    else:
        if img.mode != "RGB":
            img = img.convert("RGB")
        img.save(buf, format="JPEG", quality=92)
        mime = "image/jpeg"
    return f"data:{mime};base64,{base64.b64encode(buf.getvalue()).decode('ascii')}"


def _call_vision_api(content_parts: list) -> dict:
    """通用调用：传一个 content 列表（可能含 text + 多张 image），返回解析后的 dict。"""
    payload = {
        "model": MODEL_NAME,
        "messages": [
            {
                "role": "system",
                "content": "你是一个严格的 JSON 输出助手，只输出 JSON，不要任何额外文字。",
            },
            {
                "role": "user",
                "content": content_parts,
            },
        ],
        "temperature": 0.2,
        "max_tokens": 1500,
    }
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
    }
    resp = requests.post(
        f"{API_BASE.rstrip('/')}/chat/completions",
        headers=headers,
        json=payload,
        timeout=120,
    )
    resp.raise_for_status()
    content = resp.json()["choices"][0]["message"]["content"].strip()

    # 模型偶尔会包 ```json ... ```，清洗一下
    if content.startswith("```"):
        content = content.strip("`")
        if content.lower().startswith("json"):
            content = content[4:]
        content = content.strip()
    return json.loads(content)


def _normalize_record(record: dict) -> dict:
    """统一一个提取结果：为所有 FIELDS 补默认值；归一化 category；去时区。"""
    for f in FIELDS:
        record.setdefault(f, None)
    if record.get("category") not in CATEGORY_LIST:
        record["category"] = "lecture"
    # 去时区：避免表格里出现 "+08:00"，避免写入 Excel 时多余字串
    for _col in ("time_start", "time_end"):
        v = record.get(_col)
        if v is None or v == "":
            record[_col] = None
            continue
        try:
            ts = pd.to_datetime(v, errors="coerce")
            if pd.isna(ts):
                record[_col] = None
                continue
            # 优先转为上海时区再丢掉 tzinfo，保留字面“本地时间”
            try:
                ts = ts.tz_convert("Asia/Shanghai").tz_localize(None)
            except (AttributeError, TypeError):
                # 没 tzinfo，直接本地化
                try:
                    ts = ts.tz_localize("Asia/Shanghai").tz_localize(None)
                except Exception:
                    ts = ts.tz_localize(None)
            record[_col] = ts.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            record[_col] = None
    return record


def extract_lecture_from_image(img: Image.Image) -> dict:
    """调用 DeepSeek 大模型从图片里提取讲座/活动/场所信息。"""
    data_url = encode_image_to_data_url(img)
    content = [
        {"type": "text", "text": EXTRACTION_PROMPT},
        {"type": "image_url", "image_url": {"url": data_url}},
    ]
    record = _call_vision_api(content)
    return _normalize_record(record)


def save_uploaded_image(uploaded) -> str:
    """保存用户上传的图片，返回 base64 数据串（存进 source_image，线上可内嵌展示）。
    同时落一份到 uploads/ 目录作为本地兜底。"""
    import base64
    data = uploaded.getbuffer()
    b64 = base64.b64encode(data).decode("ascii")
    try:
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        ext = os.path.splitext(uploaded.name or "")[1].lower()
        if ext not in (".png", ".jpg", ".jpeg", ".webp"):
            ext = ".png"
        fname = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}{ext}"
        with open(os.path.join(UPLOAD_DIR, fname), "wb") as f:
            f.write(data)
    except Exception:
        pass
    return b64


# ============== 微信公众号文章抓取 ==============
WX_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Referer": "https://mp.weixin.qq.com/",
}

# 微信风控提示词，命中即认为被抓取拦截
WX_BLOCK_FLAGS = [
    "环境异常", "去验证", "参数错误",
    "此内容因违规无法查看", "已被发布者删除",
    "该内容已被发布者删除", "访问过于频繁",
]


def fetch_wechat_article(url: str) -> dict:
    """抓取微信公众号文章，返回 {title, account, text, images}。

    - text: 正文纯文本（文字型文章）
    - images: PIL.Image 列表（图片型海报文章）
    """
    url = url.strip()
    resp = requests.get(url, headers=WX_HEADERS, timeout=30)
    resp.raise_for_status()
    html = resp.text

    for flag in WX_BLOCK_FLAGS:
        if flag in html:
            raise RuntimeError(f"疑似被微信风控拦截（命中：{flag}），稍后重试或更换网络。")

    doc = LH.fromstring(html)

    # 标题
    t = doc.xpath('//h1[contains(@class,"rich_media_title")]//text()')
    title = "".join(t).strip()

    # 公众号名
    acc = doc.xpath('//*[@id="js_name"]//text()')
    account = "".join(acc).strip()

    # 正文
    nodes = doc.xpath('//*[@id="js_content"]')
    text = ""
    img_urls = []
    if nodes:
        text = re.sub(r"[ \t\u00a0]+", " ", nodes[0].text_content())
        text = re.sub(r"\n\s*\n+", "\n", text).strip()
        for im in nodes[0].xpath(".//img"):
            u = im.get("data-src") or im.get("src")
            if u and u.startswith("http"):
                img_urls.append(u)

    # 下载图片（图片型文章）
    images = []
    for u in img_urls[:6]:
        try:
            ir = requests.get(u, headers=WX_HEADERS, timeout=30)
            if ir.status_code == 200 and len(ir.content) > 3000:
                images.append(Image.open(io.BytesIO(ir.content)).convert("RGB"))
        except Exception:
            continue

    return {"title": title, "account": account, "text": text, "images": images}


def image_to_data_url(img: Image.Image) -> str:
    """把 PIL Image 转成 JPEG base64 data URL（用于链接抓取的图片）。"""
    buf = io.BytesIO()
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.save(buf, format="JPEG", quality=90)
    return f"data:image/jpeg;base64,{base64.b64encode(buf.getvalue()).decode('ascii')}"


def extract_lecture_from_article(article: dict) -> dict:
    """从抓取到的公众号文章提取讲座/活动/场所信息（文字型送文本，图片型送图片）。"""
    content_parts = [{"type": "text", "text": EXTRACTION_PROMPT}]
    if article.get("text"):
        content_parts.append({"type": "text", "text": "【文章正文】\n" + article["text"]})
    for img in article.get("images", [])[:4]:
        content_parts.append({"type": "image_url", "image_url": {"url": image_to_data_url(img)}})
    if len(content_parts) == 1:
        raise RuntimeError("该文章既无文字正文也无可用图片，无法提取。")

    record = _call_vision_api(content_parts)
    # 用抓取到的公众号名兜底
    if not record.get("source_account") and article.get("account"):
        record["source_account"] = article["account"]
    return _normalize_record(record)


def extract_from_url(url: str) -> tuple[dict, dict]:
    """完整链路：URL -> 抓取 -> 提取。返回 (record, article_meta)。"""
    article = fetch_wechat_article(url)
    record = extract_lecture_from_article(article)
    meta = {
        "url": url,
        "article_title": article["title"],
        "account": article["account"],
        "text_len": len(article["text"]),
        "image_count": len(article["images"]),
    }
    return record, meta


# ============== 业务字段定义 ==============
CATEGORY_LABELS = {
    "lecture":            "讲座信息",
    "humanities_activity": "人文活动",
    "study_venue":        "学习场所开放",
}
CATEGORY_LIST = list(CATEGORY_LABELS.keys())

# 所有可能插入的列（SQL 插入、Excel 导出都会用到）
FIELDS = [
    "category",
    "title", "speaker", "affiliation",
    "time_text", "time_start", "time_end",
    "location", "online_url",
    "host", "registration", "deadline",
    "source_account", "source_url", "source_type", "source_image", "summary",
    "theme", "participants",
    "venue", "hours", "open_period", "contact",
]

# 各分类重点展示字段（编辑表单 / 表格）
FIELDS_BY_CATEGORY = {
    "lecture": [
        "title", "speaker", "affiliation",
        "time_text", "time_start", "time_end",
        "location", "online_url",
        "host", "registration", "deadline",
        "source_account", "summary",
    ],
    "humanities_activity": [
        "title", "theme", "participants",
        "time_text", "time_start", "time_end",
        "location", "online_url",
        "host", "registration", "deadline",
        "source_account", "summary",
    ],
    "study_venue": [
        "title", "venue", "hours", "open_period",
        "host", "location", "contact",
        "source_account", "summary",
    ],
}

# 字段中文标签（极简列表视图 / 详情展示用）
FIELD_LABELS = {
    "title":        "标题",
    "speaker":      "主讲人",
    "affiliation":  "单位",
    "time_text":    "时间",
    "time_start":   "开始时间",
    "time_end":     "结束时间",
    "location":     "地点",
    "online_url":   "线上链接",
    "host":         "主办方",
    "registration": "报名方式",
    "deadline":     "截止时间",
    "source_account": "来源账号",
    "source_url":   "消息来源",
    "source_type":  "来源类型",
    "summary":      "摘要",
    "theme":        "主题",
    "participants": "参与对象",
    "venue":        "场所",
    "hours":        "开放时间",
    "open_period":  "有效时段",
    "contact":      "联系方式",
}

# ============== UI: 侧栏信息（只读） ==============
init_db()

with st.sidebar:
    st.header("系统信息")
    st.caption("模型：DeepSeek (deepseek-chat)")
    st.caption("接口：https://api.deepseek.com/v1")
    st.caption("API Key：已配置")
    db_name = TURSO_URL.replace("libsql://", "") if USE_TURSO else os.path.basename(LOCAL_DB)
    st.caption(f"数据库：{db_name}")
    _all = load_all_lectures()
    st.caption(f"当前共 {len(_all)} 条记录")
    if not _all.empty:
        _all = _all.copy()
        if "category" not in _all.columns:
            _all["category"] = "lecture"
        else:
            _all["category"] = _all["category"].fillna("lecture")
        _counts = _all["category"].value_counts().to_dict()
        _cat_summary = " ｜ ".join(
            f"{CATEGORY_LABELS[cat]} {_counts.get(cat, 0)}"
            for cat in CATEGORY_LIST
        )
        st.caption(f"分类：{_cat_summary}")

# ============== UI: 查询 ==============
st.divider()
st.subheader("信息库查询")

df = load_all_lectures()
if df.empty:
    st.caption("信息库尚为空 · 点击右上角「导入新信息」开始录入")
else:
    # category 默认为空（防御性补列，应对缓存中的旧 DataFrame）
    if "category" not in df.columns:
        df["category"] = "lecture"
    df["category"] = df["category"].fillna("lecture")
    df["_cat_label"] = df["category"].map(lambda k: CATEGORY_LABELS.get(k, "讲座信息"))

    # ---------- 分类分布统计（供搜索行右侧小字使用） ----------
    counts = df["category"].value_counts().to_dict()

    # ---------- 分类筛选（极简下划线 Tab，多选） ----------
    # 用 st.pills（selection_mode="multi"）承载多选，CSS 覆盖成下划线样式：
    # 选中黑字 + 底部下划线，未选灰字。状态仍走 cat_filter。
    st.markdown(
        """
    <style>
    /* 下划线 Tab：去掉 pills 的胶囊底色/边框/圆角，仅保留底部下划线作为选中态 */
    [data-testid="stPills"] [data-testid="stBaseButton"] {
      background: transparent !important;
      border: none !important;
      box-shadow: none !important;
      border-radius: 0 !important;
      padding: 6px 14px !important;
    }
    [data-testid="stPills"] [data-testid="stBaseButton"] p {
      color: #8A8F98 !important;
      font-weight: 400 !important;
    }
    [data-testid="stPills"] [data-testid="stBaseButton"]:hover p {
      color: #2D3748 !important;
    }
    /* 选中项：黑字 + 底部下划线 */
    [data-testid="stPills"] [data-testid="stBaseButton"][aria-selected="true"] {
      border-bottom: 2px solid #2D3748 !important;
    }
    [data-testid="stPills"] [data-testid="stBaseButton"][aria-selected="true"] p {
      color: #2D3748 !important;
      font-weight: 600 !important;
    }
    /* 选中项默认那个对勾/填充隐藏 */
    [data-testid="stPills"] [data-testid="stBaseButton"] [data-testid="stPill"] {
      background: transparent !important;
    }
    </style>
    """,
        unsafe_allow_html=True,
    )
    st.pills(
        "分类",
        options=CATEGORY_LIST,
        default=list(CATEGORY_LIST),
        selection_mode="multi",
        format_func=lambda k: CATEGORY_LABELS.get(k, k),
        key="cat_filter",
        label_visibility="collapsed",
    )

    # ---------- 关键词搜索 + 右侧记录数小字 ----------
    _kw_col, _cnt_col = st.columns([3, 1])
    keyword = _kw_col.text_input(
        "搜索", "", key="kw_all",
        placeholder="搜索讲座或活动（支持主讲人、单位、地点）",
        label_visibility="collapsed",
    )
    _kw_col.caption("例如：语言学、复旦、Volker")
    _cnt_col.write("")
    _cnt_col.caption(
        f"共 {len(df)} 条 · "
        + " · ".join(
            f"{CATEGORY_LABELS.get(cat, cat)} {counts.get(cat, 0)}"
            for cat in CATEGORY_LIST
        )
    )

    with st.expander("高级筛选（可选）", expanded=False):
        f1, f2, f3 = st.columns(3)
        kw_title   = f1.text_input("按标题",   "", key="kw_title")
        kw_speaker = f2.text_input("按主讲人", "", key="kw_speaker")
        kw_host    = f3.text_input("按主办方", "", key="kw_host")

        valid_ts = pd.to_datetime(df["time_start"], errors="coerce").dropna()
        if not valid_ts.empty:
            default_range = (valid_ts.min().date(), valid_ts.max().date())
        else:
            default_range = (datetime.now().date(), datetime.now().date())
        use_date = st.checkbox("启用时间范围筛选", value=False, key="use_date")
        date_min, date_max = st.date_input(
            "时间范围（开始 ~ 结束）",
            value=default_range,
            key="date_range",
            disabled=not use_date,
        )

    view = df.copy()

    # 分类筛选
    sel_cats = st.session_state.get("cat_filter", CATEGORY_LIST) or CATEGORY_LIST
    view = view[view["category"].isin(sel_cats)]

    # 全局关键词
    if keyword:
        search_cols = [c for c in view.columns if c not in ("id", "created_at", "_cat_label")]
        mask = pd.Series(False, index=view.index)
        for c in search_cols:
            mask = mask | view[c].fillna("").astype(str).str.contains(keyword, case=False, na=False)
        view = view[mask]

    # 高级筛选
    if kw_title:
        view = view[view["title"].fillna("").str.contains(kw_title, case=False, na=False)]
    if kw_speaker:
        view = view[view["speaker"].fillna("").str.contains(kw_speaker, case=False, na=False)]
    if kw_host:
        view = view[view["host"].fillna("").str.contains(kw_host, case=False, na=False)]

    if use_date and date_min and date_max:
        ts = pd.to_datetime(view["time_start"], errors="coerce")
        view = view[
            (ts >= pd.to_datetime(date_min))
            & (ts <= pd.to_datetime(date_max) + pd.Timedelta(days=1))
        ]

    # ---------- 按分类分组展示（极简列表视图） ----------
    _sheet_color = {
        "lecture":            "8A9AAB",
        "humanities_activity":"B4A597",
        "study_venue":        "A3B5A6",
    }
    _sheet_short = {
        "lecture":            "Lecture",
        "humanities_activity":"Activity",
        "study_venue":        "Venue",
    }
    def _fmt_time(row):
        """优先可读时间 text，否则用 time_start 的日期时间。"""
        t = row.get("time_text")
        if t is not None and not pd.isna(t) and str(t).strip() and str(t).strip() != "nan":
            return str(t).strip()
        ts = row.get("time_start")
        if ts is not None and not pd.isna(ts) and str(ts) != "" and str(ts) != "NaT":
            try:
                return pd.to_datetime(ts).strftime("%m月%d日 %H:%M")
            except Exception:
                return str(ts)
        return ""

    def _cat_subline(row, cat_key):
        """第二行灰色小字：按分类取关键字段拼接。"""
        if cat_key == "lecture":
            parts = [row.get("speaker"), row.get("affiliation")]
        elif cat_key == "humanities_activity":
            parts = [row.get("theme"), row.get("participants"), row.get("speaker")]
        else:  # study_venue
            parts = [row.get("venue"), row.get("hours")]
        out = []
        for p in parts:
            if p is None or (isinstance(p, float) and pd.isna(p)):
                continue
            s = str(p).strip()
            if s and s not in ("nan", "None", "NaT"):
                out.append(s)
        sub = " · ".join(out)
        # 主讲人/单位过长时截断（避免 17 个人名 + 10 个单位撑满第二行）
        if len(sub) > 48:
            sub = sub[:48].rstrip() + "…"
        return sub

    def _render_list_item_html(row, cat_key):
        title = str(row.get("title") or "").strip() or "（无标题）"
        subline = _cat_subline(row, cat_key)
        # 右侧：时间优先（完整），地点次之（过长可截断）
        t = _fmt_time(row)
        # 时间只取第一段（多场次排期用 ; ； , ， 分隔时只显示首场）
        t_first = t
        for _sep in ("；", ";", "，", ","):
            if _sep in t_first:
                t_first = t_first.split(_sep)[0].strip()
                break
        loc = row.get("location")
        loc_s = ""
        if loc is not None and not pd.isna(loc) and str(loc).strip() and str(loc).strip() != "nan":
            loc_s = str(loc).strip()
        # 拼接：地点 · 时间；整体超长时截地点、保时间完整
        if t_first and loc_s:
            side = f"{loc_s} · {t_first}"
        elif t_first:
            side = t_first
        elif loc_s:
            side = loc_s
        else:
            side = ""
        if len(side) > 60:
            # 尽量保留时间完整：截断地点部分
            if t_first and t_first in side:
                keep_tail = t_first
                head_budget = 60 - len(keep_tail) - 3
                if head_budget > 0:
                    side = loc_s[:head_budget].rstrip() + "…" + " · " + keep_tail
                else:
                    side = side[:60].rstrip() + "…"
            else:
                side = side[:60].rstrip() + "…"
        # 消息来源（统一作为跳转接口：链接→原文，图片→data URI 原图）
        src = row.get("source_url")
        src_type = str(row.get("source_type") or "").strip()
        src_img = row.get("source_image")
        src_txt = ""
        if src_type == "image" and src_img is not None and not pd.isna(src_img) and str(src_img).strip():
            # 图片来源：内嵌 base64，点开新标签看原图
            src_txt = (
                f'<a href="data:image/png;base64,{str(src_img).strip()}" '
                f'target="_blank" class="list-src">查看原图 ↗</a>'
            )
        elif src is not None and not pd.isna(src) and str(src).strip() and str(src).strip() not in ("nan", "None"):
            s = str(src).strip()
            if s.startswith("http"):
                src_txt = f'<a href="{s}" target="_blank" class="list-src">原文 ↗</a>'
            else:
                src_txt = f'<span class="list-src" style="color:#B5BAC1;">来源：{s}</span>'
        meta_html = ""
        if subline:
            meta_html += f'<span>{subline}</span>'
        if src_txt:
            meta_html += (f'<span style="margin-left:10px;">{src_txt}</span>' if meta_html else src_txt)
        # 注意：必须输出单行 HTML（无换行无缩进）——Markdown 会把 4+ 空格缩进的行当作代码块原样显示
        return (
            f'<div class="list-item"><div class="list-main">'
            f'<div class="list-title">{title}</div>'
            f'<div class="list-meta">{meta_html}</div>'
            f'</div><div class="list-side">{side}</div></div>'
        )

    for cat_key in CATEGORY_LIST:
        sub = view[view["category"] == cat_key].sort_values("time_start", na_position="last")
        _n = len(sub)
        # 默认展开第一个有内容的分类，其余折叠，避免页面过长
        _default_open = (cat_key == CATEGORY_LIST[0])
        with st.expander(
            f"{CATEGORY_LABELS[cat_key]} · {_n} 条",
            expanded=_default_open,
        ):
            # 导出按钮（分类内，右上角次要按钮）
            if _n > 0:
                _exp_l, _exp_r = st.columns([5, 1])
                with _exp_r:
                    _do_export = st.button(
                        "导出 Excel", key=f"export_{cat_key}",
                        use_container_width=True, type="secondary",
                    )
            else:
                _do_export = False
            if sub.empty:
                st.markdown(
                    '<div class="list-empty">暂无相关开放信息，请点击右上角导入。</div>',
                    unsafe_allow_html=True,
                )
            else:
                # 收集所有行，包进固定高度滚动容器，避免撑长页面
                _items_html = "".join(
                    _render_list_item_html(row, cat_key)
                    for _, row in sub.iterrows()
                )
                st.markdown(
                    f'<div class="list-scroll">{_items_html}</div>',
                    unsafe_allow_html=True,
                )
            # 该分类的导出
            if _do_export:
                with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
                    cols = [c for c in FIELDS if c in sub.columns]
                    sheet_name = _sheet_short.get(cat_key, cat_key)[:31]
                    sub[cols].to_excel(writer, index=False, sheet_name=sheet_name)
                    try:
                        writer.sheets[sheet_name].sheet_properties.tabColor = (
                            _sheet_color.get(cat_key, "8A9AAB")
                        )
                    except Exception:
                        pass
                with open(EXCEL_PATH, "rb") as f:
                    st.download_button(
                        "下载 lectures.xlsx",
                        data=f.read(),
                        file_name=f"lectures_{_sheet_short.get(cat_key, cat_key)}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{cat_key}",
                    )
                st.rerun()

    # 全局导出（所有分类合并，按分类多 sheet）—— 保留在底部一个次要按钮
    _exp_all_l, _exp_all_r = st.columns([5, 1])
    with _exp_all_r:
        _do_export_all = st.button(
            "导出全部", key="export_all", use_container_width=True, type="secondary",
        )
    if _do_export_all:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            written = False
            for cat_key in CATEGORY_LIST:
                sub = view[view["category"] == cat_key]
                if sub.empty:
                    continue
                written = True
                cols = [c for c in FIELDS if c in sub.columns]
                sheet_name = _sheet_short.get(cat_key, cat_key)[:31]
                sub[cols].to_excel(writer, index=False, sheet_name=sheet_name)
                try:
                    writer.sheets[sheet_name].sheet_properties.tabColor = (
                        _sheet_color.get(cat_key, "8A9AAB")
                    )
                except Exception:
                    pass
            if not written:
                view[FIELDS].to_excel(writer, index=False, sheet_name="All")
        with open(EXCEL_PATH, "rb") as f:
            st.download_button(
                "下载 lectures.xlsx",
                data=f.read(),
                file_name="lectures.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_all",
            )
        st.rerun()

    # 扁平表（备选查看）—— 纯按时间顺序（早→晚）
    with st.expander("扁平表（所有记录按时间顺序）", expanded=False):
        flat_view = view.sort_values(["time_start", "category"], na_position="last")
        flat_cols = ["_cat_label"] + [c for c in FIELDS if c in flat_view.columns and c != "category"]
        _flat = flat_view[flat_cols].copy()
        _flat = _flat.rename(columns={"_cat_label": "分类", **FIELD_LABELS})
        st.dataframe(_flat, width='stretch', hide_index=True)

@st.dialog("导入新信息", width="large")
def render_import():
    input_mode = st.radio(
        "选择输入方式",
        ["上传截图", "微信公众号链接"],
        horizontal=True,
        key="input_mode",
    )

    # ---------- 方式一：上传截图 ----------
    if input_mode == "上传截图":
        uploaded = st.file_uploader(
            "拖拽 / 选择一张讲座海报或公众号文章截图",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=False,
            key="upload_single",
        )

        col1, col2 = st.columns([1, 2])
        with col1:
            if uploaded is not None:
                img = Image.open(uploaded)
                st.image(img, caption="已上传的截图", width='stretch')

        with col2:
            if uploaded is not None and st.button("提取信息", type="primary"):
                with st.spinner("正在调用大模型提取，请稍候..."):
                    try:
                        record = extract_lecture_from_image(img)
                        record["source_image"] = save_uploaded_image(uploaded)
                        record["source_type"] = "image"
                        st.session_state["pending_record"] = record
                        label = CATEGORY_LABELS.get(record.get("category"), "讲座信息")
                        st.success(f"识别为：{label}，请在下方确认后写入。")
                    except Exception as e:
                        st.error(f"提取失败: {e}")

    # ---------- 方式二：微信公众号链接（批量） ----------
    else:
        st.caption("支持一次粘贴多条链接（每行一条），常见格式：https://mp.weixin.qq.com/s/xxxxx")
        url_text = st.text_area(
            "微信公众号文章链接（每行一条）",
            height=160,
            placeholder="https://mp.weixin.qq.com/s/xxxxxxxxxx\nhttps://mp.weixin.qq.com/s/yyyyyyyyyy",
            key="wx_urls",
        )
        only_image = st.checkbox("仅保留成功提取到标题的记录", value=True, key="wx_filter")

        if st.button("开始批量抓取并提取", type="primary"):
            urls = [u.strip() for u in url_text.splitlines() if u.strip()]
            urls = list(dict.fromkeys(urls))  # 去重保持顺序
            if not urls:
                st.warning("请先粘贴至少一条链接。")
            else:
                results, errors = [], []
                bar = st.progress(0.0, text="开始抓取...")
                for i, u in enumerate(urls, 1):
                    bar.progress((i - 1) / len(urls), text=f"[{i}/{len(urls)}] 处理中：{u[:60]}...")
                    try:
                        rec, meta = extract_from_url(u)
                        rec["_meta"] = meta
                        results.append(rec)
                    except Exception as e:
                        errors.append({"url": u, "error": str(e)})
                    time.sleep(1.5)  # 对微信友好，降低风控风险
                bar.progress(1.0, text="完成")
                st.session_state["wx_results"] = results
                st.session_state["wx_errors"] = errors
                st.session_state.pop("wx_done", None)

        # 展示抓取结果
        if st.session_state.get("wx_results") is not None:
            results = st.session_state["wx_results"]
            errors = st.session_state.get("wx_errors", [])
            st.success(f"本次共提取 {len(results)} 条，失败 {len(errors)} 条。")
            if errors:
                with st.expander(f"失败详情（{len(errors)} 条）"):
                    for e in errors:
                        st.write(f"- `{e['url'][:70]}` → {e['error']}")
            if results:
                st.markdown("**预览（可核对后再批量入库）**")
                preview_rows = []
                for r in results:
                    cat = r.get("category") if r.get("category") in CATEGORY_LIST else "lecture"
                    cols = FIELDS_BY_CATEGORY[cat]
                    preview_rows.append({
                        "分类": CATEGORY_LABELS.get(cat, cat),
                        **{k: r.get(k) for k in cols if k != "summary"}
                    })
                st.dataframe(pd.DataFrame(preview_rows), width='stretch', hide_index=True)

                if st.button("全部写入数据库", key="wx_bulk_insert"):
                    ok_n = dup_n = 0
                    for r in results:
                        meta = r.get("_meta") or {}
                        r = {k: v for k, v in r.items() if k != "_meta"}
                        if meta.get("url"):
                            r["source_url"] = meta["url"]
                        r["source_type"] = "wechat"
                        if only_image and not (r.get("title") or "").strip():
                            continue
                        ok, _ = insert_lecture(r)
                        ok_n += 1 if ok else 0
                        dup_n += 0 if ok else 1
                    st.cache_data.clear()
                    st.success(f"新写入 {ok_n} 条，跳过重复/空记录 {dup_n} 条。")
                    st.session_state.pop("wx_results", None)
                    st.session_state.pop("wx_errors", None)
                    st.rerun()

    # 确认 & 写入
    if "pending_record" in st.session_state:
        st.subheader("确认信息并写入")
        rec = st.session_state["pending_record"]
        with st.form("confirm_form"):
            # ---- 1. 分类选择（可手动调整）----
            cur_cat = rec.get("category") if rec.get("category") in CATEGORY_LIST else "lecture"
            cat_label_to_key = {v: k for k, v in CATEGORY_LABELS.items()}
            rec["category"] = st.selectbox(
                "分类（可手动调整）",
                options=CATEGORY_LIST,
                format_func=lambda k: CATEGORY_LABELS.get(k, k),
                index=CATEGORY_LIST.index(cur_cat),
            )
            st.write("")

            fields = FIELDS_BY_CATEGORY[rec["category"]]

            # ---- 2. 动态字段展示 ----
            # 先根据 category 生成不同行
            if rec["category"] == "lecture":
                c1, c2, c3 = st.columns(3)
                rec["title"]       = c1.text_input("讲座标题",  rec.get("title") or "")
                rec["speaker"]     = c2.text_input("主讲人",    rec.get("speaker") or "")
                rec["affiliation"] = c3.text_input("主讲人单位", rec.get("affiliation") or "")
            elif rec["category"] == "humanities_activity":
                c1, c2, c3 = st.columns(3)
                rec["title"]        = c1.text_input("活动名称",        rec.get("title") or "")
                rec["theme"]        = c2.text_input("主题 / 征稿主题", rec.get("theme") or "")
                rec["participants"] = c3.text_input("参与对象 / 参赛要求", rec.get("participants") or "")
            else:  # study_venue
                c1, c2, c3 = st.columns(3)
                rec["title"]    = c1.text_input("信息标题", rec.get("title") or "")
                rec["venue"]    = c2.text_input("场所名称", rec.get("venue") or "")
                rec["hours"]    = c3.text_input("开放时间描述", rec.get("hours") or "")

            # time_text / time_start / time_end —— 讲座和人文活动用
            if rec["category"] in ("lecture", "humanities_activity"):
                c4, c5, c6 = st.columns(3)
                rec["time_text"]  = c4.text_input("原文时间描述",  rec.get("time_text") or "")
                rec["time_start"] = c5.text_input("开始时间 ISO8601", rec.get("time_start") or "")
                rec["time_end"]   = c6.text_input("结束时间 ISO8601", rec.get("time_end") or "")

            # 场地 / 开放期 —— 场所
            if rec["category"] == "study_venue":
                c4, c5 = st.columns(2)
                rec["open_period"] = c4.text_input("有效起止日期", rec.get("open_period") or "")
                rec["contact"]     = c5.text_input("联系方式（电话 / 联系人）", rec.get("contact") or "")
                # 场所也可有具体地点（推荐 location）
                rec["location"] = st.text_input("具体位置（选填）", rec.get("location") or "")

            # location / online_url —— 三类都可能有
            if rec["category"] in ("lecture", "humanities_activity"):
                c7, c8 = st.columns(2)
                rec["location"]   = c7.text_input("线下地点", rec.get("location") or "")
                rec["online_url"] = c8.text_input("线上链接", rec.get("online_url") or "")

            # host / registration / deadline —— 讲座和活动
            if rec["category"] in ("lecture", "humanities_activity"):
                c9, c10, c11 = st.columns(3)
                rec["host"]         = c9.text_input("主办方",       rec.get("host") or "")
                rec["registration"] = c10.text_input("报名方式",     rec.get("registration") or "")
                rec["deadline"]     = c11.text_input("报名 / 征稿截止", rec.get("deadline") or "")
            else:
                rec["host"] = st.text_input("主办 / 负责单位", rec.get("host") or "")

            c12, c13 = st.columns(2)
            rec["source_account"] = c12.text_input("公众号名称", rec.get("source_account") or "")
            rec["source_url"]     = c13.text_input("消息来源链接", rec.get("source_url") or "")
            rec["summary"]        = st.text_area("一句话摘要", rec.get("summary") or "")

            cc1, cc2 = st.columns(2)
            submitted = cc1.form_submit_button("写入数据库")
            cancelled = cc2.form_submit_button("取消")

            if submitted:
                ok, msg = insert_lecture(rec)
                if ok:
                    cat_zh = CATEGORY_LABELS.get(rec["category"], "")
                    st.success(f"已写入【{cat_zh}】《{rec['title']}》。")
                    st.session_state.pop("pending_record", None)
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.warning(msg)
            if cancelled:
                st.session_state.pop("pending_record", None)
                st.rerun()


@st.dialog("录入口令")
def render_passcode():
    st.caption("信息库可自由浏览；录入新信息需要访客口令。")
    pwd = st.text_input("访客口令", type="password", key="guest_pwd")
    if st.button("验证", type="primary", key="auth_btn"):
        if not GUEST_PASSCODE:
            st.error("服务端未配置口令，请联系管理员。")
        elif pwd == GUEST_PASSCODE:
            st.session_state["auth_ok"] = True
            st.session_state["_req_import"] = True
            st.rerun(scope="app")
        else:
            st.error("口令错误。")


if st.session_state.pop("_req_pass", False):
    render_passcode()
if st.session_state.pop("_req_import", False) and st.session_state.get("auth_ok"):
    render_import()
