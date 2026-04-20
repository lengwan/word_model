"""
论文格式一键体检 - Streamlit Web 应用
运行: streamlit run app.py
"""
import streamlit as st
import tempfile
import os
import json
import hashlib
import time
import copy
import string
import random
import html as html_mod
import subprocess
from datetime import datetime
from thesis_checker import ThesisChecker, DEFAULT_RULES, merge_rules
try:
    from thesis_checker import get_default_rules
except ImportError:
    # 向后兼容：若线上 thesis_checker.py 是旧版（无学历切换），退化为硕士规则
    def get_default_rules(degree='硕士'):
        import copy as _copy
        return _copy.deepcopy(DEFAULT_RULES)

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(
    page_title="论文格式一键体检",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="auto",
)

# ============================================================
# 全局 CSS
# ============================================================
st.markdown("""
<style>
/* ============================================================
   CSS 变量 - 自动适配亮/暗主题
   ============================================================ */
:root {
    --bg-primary: #0c1222;
    --bg-card: rgba(26,35,50,0.85);
    --bg-card-hover: rgba(30,42,60,0.95);
    --border-card: rgba(255,255,255,0.08);
    --text-primary: #f1f5f9;
    --text-secondary: #94a3b8;
    --text-muted: #64748b;
    --accent-blue: #3b82f6;
    --accent-indigo: #6366f1;
    --accent-green: #22c55e;
    --accent-red: #ef4444;
    --accent-yellow: #eab308;
    --tag-bg: rgba(59,130,246,0.15);
    --tag-border: rgba(59,130,246,0.35);
    --tag-text: #60a5fa;
    --metric-bg: rgba(26,35,50,0.85);
    --metric-label: #94a3b8;
    --paywall-gradient: rgba(12,18,34,0.95);
    --divider: rgba(255,255,255,0.08);
    --score-card-bg: rgba(30,42,60,0.9);
}

/* 亮色模式覆盖 */
@media (prefers-color-scheme: light) {
    :root {
        --bg-primary: #ffffff;
        --bg-card: #f8fafc;
        --bg-card-hover: #f1f5f9;
        --border-card: rgba(0,0,0,0.08);
        --text-primary: #1e293b;
        --text-secondary: #475569;
        --text-muted: #94a3b8;
        --accent-blue: #2563eb;
        --accent-indigo: #4f46e5;
        --tag-bg: rgba(37,99,235,0.08);
        --tag-border: rgba(37,99,235,0.25);
        --tag-text: #2563eb;
        --metric-bg: #f1f5f9;
        --metric-label: #475569;
        --paywall-gradient: rgba(255,255,255,0.95);
        --divider: rgba(0,0,0,0.06);
        --score-card-bg: #f1f5f9;
    }
}

/* Streamlit 白色背景兜底 */
[data-testid="stAppViewContainer"][style*="background-color: rgb(255"] {
    --bg-card: #f8fafc;
    --bg-card-hover: #f1f5f9;
    --border-card: rgba(0,0,0,0.08);
    --text-primary: #1e293b;
    --text-secondary: #475569;
    --text-muted: #94a3b8;
    --tag-bg: rgba(37,99,235,0.08);
    --tag-border: rgba(37,99,235,0.25);
    --tag-text: #2563eb;
    --metric-bg: #f1f5f9;
    --metric-label: #475569;
    --score-card-bg: #f1f5f9;
}

/* ---- 顶部渐变条 ---- */
div[data-testid="stAppViewContainer"]::before {
    content:''; display:block; height:3px;
    background:linear-gradient(90deg, var(--accent-blue), var(--accent-indigo));
    position:fixed; top:0; left:0; right:0; z-index:9999;
}

/* ---- 隐藏默认元素 ---- */
header[data-testid="stHeader"] { background:transparent; }
footer { visibility:hidden; }
.stDeployButton, [data-testid="stToolbar"],
div[data-testid="stDecoration"] { display:none!important; }
#MainMenu { visibility:hidden; }
[data-testid="stToast"] { display:none!important; }

/* ---- 全局字体 ---- */
html, body, [class*="css"] {
    font-family: 'PingFang SC','Microsoft YaHei','Noto Sans SC',system-ui,sans-serif;
}

/* ---- Hero ---- */
.hero { text-align:center; padding:48px 20px 32px; }
.hero h1 { font-size:2.4rem; font-weight:800; margin:0 0 8px; color:var(--text-primary); }
.hero p { font-size:1.05rem; color:var(--text-secondary); max-width:600px; margin:0 auto; }

/* ---- 数字亮点 ---- */
.highlights { display:flex; justify-content:center; gap:48px; margin:28px 0 12px; flex-wrap:wrap; }
.hl-item { text-align:center; padding:12px 16px; border-radius:12px; }
.hl-num { font-size:2.2rem; font-weight:800; line-height:1; color:var(--accent-blue); }
.hl-label { font-size:0.78rem; color:var(--text-muted); margin-top:6px; }

/* ---- 上传区 ---- */
.upload-zone {
    max-width:720px; margin:0 auto 8px;
    background:var(--bg-card);
    border:2px solid var(--tag-border);
    border-radius:16px; padding:20px 24px;
    transition: border-color 0.3s, box-shadow 0.3s;
}
.upload-zone:hover {
    border-color:var(--accent-blue);
    box-shadow:0 0 24px rgba(59,130,246,0.15);
}
.upload-zone::before {
    content:"上传论文，立即检测"; display:block; text-align:center;
    font-size:1.05rem; font-weight:700; color:var(--text-primary);
    margin-bottom:12px; padding-bottom:12px;
    border-bottom:1px solid var(--divider);
}

/* ---- 汉化上传组件 ---- */
[data-testid="stFileUploaderDropzone"] [data-testid="stMarkdownContainer"] p {
    font-size:0!important; line-height:0!important;
}
[data-testid="stFileUploaderDropzone"] [data-testid="stMarkdownContainer"] p::after {
    content:"拖放论文文件到此处"; font-size:0.95rem; color:var(--text-secondary);
}
[data-testid="stFileUploaderDropzone"] small { font-size:0!important; }
[data-testid="stFileUploaderDropzone"] small::after {
    content:"支持 .doc / .docx 格式，最大 200MB"; font-size:0.75rem; color:var(--text-muted);
}
[data-testid="stFileUploaderDropzone"] button {
    font-size:0!important; min-height:42px; padding:0 24px!important;
    background:linear-gradient(135deg,#3b82f6,#6366f1)!important;
    border:none!important; border-radius:8px!important; color:white!important;
}
[data-testid="stFileUploaderDropzone"] button:hover {
    transform:translateY(-1px); box-shadow:0 4px 16px rgba(59,130,246,0.4);
}
[data-testid="stFileUploaderDropzone"] button::after {
    content:"选择论文文件"; font-size:0.9rem; font-weight:600;
}

/* ---- 通用卡片 ---- */
.glass-card {
    background:var(--bg-card);
    border:1px solid var(--border-card);
    border-radius:12px; padding:20px; margin-bottom:12px;
    transition: transform 0.2s, box-shadow 0.2s;
}
.glass-card:hover {
    transform:translateY(-2px);
    background:var(--bg-card-hover);
    box-shadow:0 4px 16px rgba(0,0,0,0.15);
}

/* 模块卡片左边条 */
.mod-grid .glass-card { border-left:3px solid var(--accent-blue); }

/* ---- Metric 卡片 ---- */
div[data-testid="stMetric"] {
    background:var(--metric-bg);
    border:1px solid var(--border-card);
    border-radius:12px; padding:14px;
}
div[data-testid="stMetric"] label {
    font-size:0.85rem!important; color:var(--metric-label)!important; font-weight:600!important;
}
div[data-testid="stMetric"] [data-testid="stMetricValue"] {
    font-size:1.8rem!important; color:var(--text-primary)!important; font-weight:700!important;
}

/* ---- 问题卡片 ---- */
.issue-card {
    padding:14px 18px; margin-bottom:8px;
    background:var(--bg-card);
    border:1px solid var(--border-card);
    border-radius:0 10px 10px 0;
}

/* ---- 付费墙 ---- */
.paywall {
    background:linear-gradient(180deg,transparent,var(--paywall-gradient) 50%);
    padding:80px 20px 40px; text-align:center; border-radius:12px;
    margin-top:-60px; position:relative; z-index:10;
}

/* ---- 模块标签（胶囊样式）---- */
.module-tag {
    display:inline-block; padding:6px 16px; margin:4px;
    border-radius:20px; font-size:0.85rem; font-weight:500;
    background:var(--tag-bg);
    border:1px solid var(--tag-border);
    color:var(--tag-text);
}

/* ---- 套餐卡片 ---- */
.tier-free {
    border:1px solid var(--border-card);
    background:var(--bg-card);
}
.tier-free .price { color:var(--text-muted); }

.tier-basic {
    border:2px solid rgba(59,130,246,0.4);
    background:var(--bg-card);
}
.tier-basic .price { color:var(--accent-blue); }

.tier-pro {
    border:2px solid rgba(99,102,241,0.6);
    background:var(--bg-card);
    box-shadow:0 0 24px rgba(99,102,241,0.12);
}
.tier-pro .price { color:var(--accent-indigo); }

.original-price { color:var(--text-muted); text-decoration:line-through; font-size:0.85rem; }
.discount-badge {
    display:inline-block; padding:2px 10px; border-radius:4px;
    font-size:0.75rem; font-weight:600;
    background:rgba(99,102,241,0.2); color:var(--accent-indigo);
}
.recommend-badge {
    display:inline-block; padding:2px 10px; border-radius:4px;
    font-size:0.75rem; font-weight:600;
    background:rgba(239,68,68,0.15); color:var(--accent-red);
}
.tier-feature { color:var(--text-secondary); font-size:0.85rem; line-height:1.8; }

/* ---- 网格 ---- */
.pricing-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:16px; margin-bottom:12px; }
.mod-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; }

/* ---- 环形图 ---- */
.score-ring { text-align:center; }

/* ---- 分隔线 ---- */
.divider { height:1px; margin:28px 0; background:linear-gradient(90deg,transparent,var(--divider),transparent); }

/* ---- 页脚 ---- */
.app-footer {
    text-align:center; color:var(--text-muted); font-size:0.75rem;
    padding:32px 16px 16px; border-top:1px solid var(--divider); margin-top:40px;
}

/* ---- 全局按钮 ---- */
button[kind="primary"], .stButton > button[data-testid="stBaseButton-primary"] {
    background:linear-gradient(135deg,#3b82f6,#6366f1)!important;
    border:none!important; border-radius:10px!important; font-weight:600!important;
    color:white!important;
}

/* ---- 手机端 ---- */
@media(max-width:768px) {
    .pricing-grid { grid-template-columns:1fr!important; }
    .mod-grid { grid-template-columns:repeat(2,1fr)!important; }
    .hero { padding:32px 16px 20px; }
    .hero h1 { font-size:1.75rem; }
    .highlights { gap:24px 32px; }
    .upload-zone { padding:12px 16px; margin:0 8px 8px; }
    .paywall { padding:48px 16px 32px; margin-top:-40px; }
}
@media(max-width:380px) {
    .hero h1 { font-size:1.5rem; }
    .hl-num { font-size:1.5rem; }
    .mod-grid { grid-template-columns:1fr!important; }
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# 套餐权限配置
# ============================================================
TIER_CONFIG = {
    'lite':   {'recheck_limit': 0,  'issue_limit': 5,  'show_suggestion': False, 'show_filter': False, 'can_download': False, 'rules_view': False, 'rules_edit': False, 'auto_fix': False},
    'basic':  {'recheck_limit': 3,  'issue_limit': -1, 'show_suggestion': False, 'show_filter': True,  'can_download': True,  'rules_view': True,  'rules_edit': False, 'auto_fix': False},
    'pro':    {'recheck_limit': -1, 'issue_limit': -1, 'show_suggestion': True,  'show_filter': True,  'can_download': True,  'rules_view': True,  'rules_edit': True,  'auto_fix': False},
    'fix':    {'recheck_limit': -1, 'issue_limit': -1, 'show_suggestion': True,  'show_filter': True,  'can_download': True,  'rules_view': True,  'rules_edit': True,  'auto_fix': True},
}

def _get_tier_config():
    """获取当前用户的套餐权限配置"""
    tier = st.session_state.get('user_tier', 'basic')
    return TIER_CONFIG.get(tier, TIER_CONFIG['basic'])

# ============================================================
# 兑换码管理（SQLite + 原子事务 + 会话绑定）
#
# 迁移说明：原来用 codes.json + 文件锁，存在两个问题：
#   1. Streamlit Cloud 实例重启会丢失 codes.json（临时文件系统）
#   2. TOCTOU 竞争：兑换码泄露给同学时，多个浏览器可同时尝试使用
# 改用 SQLite：
#   - WAL 模式支持并发读写
#   - BEGIN IMMEDIATE + UPDATE ... WHERE used=0 保证原子抢占
#   - 启动时自动迁移旧 codes.json
# TODO(持久化): Streamlit Cloud 免费版文件系统仍为临时，真正持久化需接
#   Supabase / Neon / GitHub Gist API。接口已抽象成 load_codes/save_codes/
#   verify_code/generate_codes，切换只需替换实现。
# ============================================================
import sqlite3

# codes.db 路径：Streamlit Cloud 上 __file__ 所在目录通常是 /mount/src/ 只读挂载，
# 落在那里会在冷启动 sqlite3.connect 就报 "unable to open database file"。
# 优先级：环境变量 CODES_DB_PATH > 系统临时目录 > 源码目录（本地开发兜底）
_CODES_DB_ENV = os.environ.get('CODES_DB_PATH')
if _CODES_DB_ENV:
    CODES_DB = _CODES_DB_ENV
else:
    _src_dir = os.path.dirname(__file__)
    _tmp_candidate = os.path.join(_src_dir, 'codes.db')
    # 尝试在源码目录创建文件，失败则落到系统临时目录
    try:
        with open(_tmp_candidate, 'a'):
            pass
        CODES_DB = _tmp_candidate
    except (OSError, PermissionError):
        CODES_DB = os.path.join(tempfile.gettempdir(), 'fhy_word_codes.db')

CODES_FILE = os.path.join(os.path.dirname(__file__), 'codes.json')  # legacy
ADMIN_PWD = "8811925123Aa!"

import uuid

def _get_session_id():
    """每个浏览器会话生成唯一 ID（存在 session_state 中，刷新不变）"""
    if 'session_id' not in st.session_state:
        st.session_state['session_id'] = str(uuid.uuid4())[:8]
    return st.session_state['session_id']

def _db():
    """打开 SQLite 连接，启用 WAL 并设置超时"""
    conn = sqlite3.connect(CODES_DB, timeout=10, isolation_level=None)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA synchronous=NORMAL')
    conn.row_factory = sqlite3.Row
    return conn

def _init_codes_db():
    """初始化表结构；若存在旧版 codes.json 则一次性迁移并备份"""
    conn = _db()
    try:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS codes (
                code TEXT PRIMARY KEY,
                tier TEXT,
                used INTEGER DEFAULT 0,
                created TEXT,
                used_at TEXT,
                session TEXT,
                report_id TEXT,
                filename TEXT
            )
        ''')
        # 迁移旧 codes.json（只做一次，迁移后改名为 .migrated 防止重复）
        if os.path.exists(CODES_FILE):
            try:
                with open(CODES_FILE, 'r', encoding='utf-8') as f:
                    legacy = json.load(f)
                for code, info in legacy.items():
                    conn.execute('''INSERT OR IGNORE INTO codes
                        (code, tier, used, created, used_at, session, report_id, filename)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)''', (
                        code,
                        info.get('tier', 'basic'),
                        1 if info.get('used') else 0,
                        info.get('created'),
                        info.get('used_at'),
                        info.get('session'),
                        info.get('report_id'),
                        info.get('filename'),
                    ))
                # 用 os.replace 而非 rename，避免 Windows 多进程下 .migrated 已存在时抛错
                os.replace(CODES_FILE, CODES_FILE + '.migrated')
            except Exception:
                # 迁移失败不致命，保留 codes.json 供下次重试
                pass
    finally:
        conn.close()

_init_codes_db()

def load_codes():
    """返回 dict 形式的全部兑换码，兼容旧接口"""
    conn = _db()
    try:
        rows = conn.execute('SELECT * FROM codes').fetchall()
        return {
            r['code']: {
                'tier': r['tier'],
                'used': bool(r['used']),
                'created': r['created'],
                'used_at': r['used_at'],
                'session': r['session'],
                'report_id': r['report_id'],
                'filename': r['filename'],
            } for r in rows
        }
    finally:
        conn.close()

def save_codes(codes):
    """逐条 upsert，保留签名以兼容现有调用。
    注意：不再做 DELETE-then-INSERT 的整体覆盖，避免与 verify_code / generate_codes
    的并发操作之间丢数据（worker B 的旧快照保存会覆盖 worker A 刚写入的新码）。"""
    conn = _db()
    try:
        conn.execute('BEGIN')
        for code, info in codes.items():
            conn.execute('''INSERT OR REPLACE INTO codes
                (code, tier, used, created, used_at, session, report_id, filename)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)''', (
                code,
                info.get('tier', 'basic'),
                1 if info.get('used') else 0,
                info.get('created'),
                info.get('used_at'),
                info.get('session'),
                info.get('report_id'),
                info.get('filename'),
            ))
        conn.execute('COMMIT')
    except Exception:
        try: conn.execute('ROLLBACK')
        except Exception: pass
        raise
    finally:
        conn.close()

def verify_code(code, report_id=None, filename=None):
    """原子性兑换：BEGIN IMMEDIATE + UPDATE WHERE used=0，防止同码多人抢用"""
    code = code.strip().upper()
    conn = _db()
    try:
        conn.execute('BEGIN IMMEDIATE')
        row = conn.execute('SELECT used FROM codes WHERE code = ?', (code,)).fetchone()
        if row is None:
            conn.execute('ROLLBACK')
            return False, '兑换码无效'
        if row['used']:
            conn.execute('ROLLBACK')
            return False, '此兑换码已被使用'
        # 原子更新：UPDATE 返回影响行数，若为 0 说明被别人抢先
        cur = conn.execute('''UPDATE codes
            SET used = 1, used_at = ?, session = ?, report_id = ?, filename = ?
            WHERE code = ? AND used = 0''', (
            datetime.now().isoformat(),
            _get_session_id(),
            report_id,
            filename,
            code,
        ))
        if cur.rowcount == 0:
            conn.execute('ROLLBACK')
            return False, '此兑换码已被使用'
        conn.execute('COMMIT')
        return True, '解锁成功'
    except Exception as e:
        try: conn.execute('ROLLBACK')
        except Exception: pass
        return False, f'验证失败：{e}'
    finally:
        conn.close()

def load_codes_safe():
    """保留原签名，SQLite 读取本身一致，直接复用 load_codes"""
    return load_codes()

def generate_codes(n=20, tier='basic'):
    """批量生成兑换码（码重复时自动跳过）"""
    new_codes = []
    conn = _db()
    try:
        for _ in range(n):
            code = f"FMT-{''.join(random.choices(string.ascii_uppercase + string.digits, k=4))}-{''.join(random.choices(string.ascii_uppercase + string.digits, k=4))}"
            try:
                conn.execute('''INSERT INTO codes (code, tier, used, created)
                    VALUES (?, ?, 0, ?)''',
                    (code, tier, datetime.now().isoformat()))
                new_codes.append(code)
            except sqlite3.IntegrityError:
                pass  # 极小概率撞码，跳过
    finally:
        conn.close()
    return new_codes

# ============================================================
# SVG 环形评分
# ============================================================
def render_score_ring(score, max_score, grade):
    pct = score / max_score * 100 if max_score > 0 else 0
    r = 70
    circ = 2 * 3.14159 * r
    offset = circ - (circ * pct / 100)
    if pct >= 80:   c1, c2 = '#10b981', '#34d399'
    elif pct >= 60: c1, c2 = '#3b82f6', '#60a5fa'
    elif pct >= 40: c1, c2 = '#f59e0b', '#fbbf24'
    else:           c1, c2 = '#ef4444', '#f87171'
    gl = {'A':'优秀','B':'良好','C':'中等','D':'及格','F':'不及格'}
    return f'''<div class="score-ring">
    <svg width="200" height="200" viewBox="0 0 200 200">
      <defs><linearGradient id="sg" x1="0%" y1="0%" x2="100%" y2="0%">
        <stop offset="0%" style="stop-color:{c1}"/><stop offset="100%" style="stop-color:{c2}"/>
      </linearGradient></defs>
      <circle cx="100" cy="100" r="{r}" fill="none" stroke="rgba(255,255,255,0.08)" stroke-width="10"/>
      <circle cx="100" cy="100" r="{r}" fill="none" stroke="url(#sg)" stroke-width="10"
        stroke-linecap="round" stroke-dasharray="{circ}" stroke-dashoffset="{offset}"
        transform="rotate(-90 100 100)"
        style="filter:drop-shadow(0 0 8px {c1}40);transition:stroke-dashoffset 1.2s ease-out;"/>
      <text x="100" y="88" text-anchor="middle" fill="#f1f5f9" font-size="42" font-weight="800">{score:.0f}</text>
      <text x="100" y="110" text-anchor="middle" fill="#64748b" font-size="14">/ {max_score}</text>
      <text x="100" y="138" text-anchor="middle" fill="{c1}" font-size="14" font-weight="700">{grade} {gl.get(grade,'')}</text>
    </svg></div>'''

# ============================================================
# 渲染问题卡片（复用）
# ============================================================
def render_issue(issue, show_suggestion=True):
    _e = html_mod.escape
    sev_c = {'error':'#f87171','warning':'#fbbf24','info':'#60a5fa'}
    src_c = {'official':'#a78bfa','supplement':'#2dd4bf','annotation':'#fb923c'}
    bc = sev_c.get(issue['severity'],'#64748b')
    sc = src_c.get(issue['source'],'#2dd4bf')
    preview = ''
    if issue.get('text_preview') and issue['text_preview'] != '(空)':
        preview = f'<div style="font-size:0.75rem;color:#64748b;margin-top:4px;">{_e(issue["text_preview"])}</div>'
    suggestion_html = ''
    if show_suggestion:
        suggestion_html = f'''<div style="font-size:0.8rem;">
        <span style="color:#10b981;">期望: {_e(issue['expected'])}</span> &nbsp;→&nbsp;
        <span style="color:#ef4444;">实际: {_e(issue['actual'])}</span>
      </div>'''
    else:
        suggestion_html = f'''<div style="font-size:0.8rem;">
        <span style="color:#ef4444;">实际: {_e(issue['actual'])}</span>
        <span style="color:#64748b;font-size:0.75rem;margin-left:8px;">🔒 升级专业版查看修改建议</span>
      </div>'''
    return f'''<div class="issue-card" style="border-left:3px solid {bc};">
      <div style="display:flex;gap:8px;margin-bottom:6px;flex-wrap:wrap;">
        <span style="background:{bc}22;color:{bc};padding:2px 10px;border-radius:4px;font-size:0.75rem;font-weight:600;">{_e(issue['severity_label'])}</span>
        <span style="background:#33415522;color:#94a3b8;padding:2px 10px;border-radius:4px;font-size:0.75rem;">{_e(issue['module'])}</span>
        <span style="background:{sc}22;color:{sc};padding:2px 10px;border-radius:4px;font-size:0.75rem;font-weight:600;">{_e(issue['source_label'])}</span>
        <span style="color:#64748b;font-size:0.75rem;font-family:monospace;">{_e(issue['location'])}</span>
      </div>
      <div style="font-size:0.9rem;color:#e2e8f0;margin-bottom:4px;">{_e(issue['rule'])}</div>
      {suggestion_html}{preview}
    </div>'''

# ============================================================
# 渲染模块卡片
# ============================================================
def render_module_card(mod, locked=False):
    """渲染单个模块卡片。locked=True时隐藏具体分数，只显示模块名和锁"""
    if locked:
        # 锁定版：灰色环 + 问号，不泄露任何分数信息
        r, sw = 22, 4
        circ = 2 * 3.14159 * r
        ring = f'''<svg width="52" height="52" viewBox="0 0 52 52" style="flex-shrink:0;">
          <circle cx="26" cy="26" r="{r}" fill="none" stroke="rgba(255,255,255,0.06)" stroke-width="{sw}"/>
          <text x="26" y="30" text-anchor="middle" fill="#64748b" font-size="14" font-weight="800">?</text>
        </svg>'''
        return f'''<div class="glass-card" style="padding:14px 16px;opacity:0.6;">
          <div style="display:flex;align-items:center;gap:12px;">
            {ring}
            <div style="flex:1;min-width:0;">
              <div style="font-weight:600;font-size:0.88rem;margin-bottom:4px;">{mod['name']}</div>
              <div style="width:100%;height:8px;background:rgba(255,255,255,0.06);border-radius:4px;overflow:hidden;">
                <div style="width:0%;height:100%;border-radius:4px;"></div>
              </div>
              <div style="font-size:0.7rem;color:#64748b;margin-top:4px;">解锁查看</div>
            </div>
          </div>
        </div>'''

    pct = mod['pct']
    if pct >= 90:   color, emoji = '#10b981', ''
    elif pct >= 70: color, emoji = '#3b82f6', ''
    elif pct >= 40: color, emoji = '#f59e0b', ''
    elif pct > 0:   color, emoji = '#ef4444', ''
    else:           color, emoji = '#ef4444', ''
    # 小型 SVG 环形百分比
    r, sw = 22, 4
    circ = 2 * 3.14159 * r
    off = circ - (circ * pct / 100)
    ring = f'''<svg width="52" height="52" viewBox="0 0 52 52" style="flex-shrink:0;">
      <circle cx="26" cy="26" r="{r}" fill="none" stroke="rgba(255,255,255,0.06)" stroke-width="{sw}"/>
      <circle cx="26" cy="26" r="{r}" fill="none" stroke="{color}" stroke-width="{sw}"
        stroke-linecap="round" stroke-dasharray="{circ}" stroke-dashoffset="{off}"
        transform="rotate(-90 26 26)"/>
      <text x="26" y="30" text-anchor="middle" fill="{color}" font-size="12" font-weight="800">{pct:.0f}</text>
    </svg>'''
    err_txt = f'<span style="color:#f87171;">{mod["errors"]}错误</span>' if mod['errors'] else ''
    warn_txt = f'<span style="color:#fbbf24;">{mod["warnings"]}警告</span>' if mod['warnings'] else ''
    sep = ' ' if err_txt and warn_txt else ''
    status = err_txt + sep + warn_txt if (err_txt or warn_txt) else '<span style="color:#10b981;">通过</span>'
    return f'''<div class="glass-card" style="padding:14px 16px;">
      <div style="display:flex;align-items:center;gap:12px;">
        {ring}
        <div style="flex:1;min-width:0;">
          <div style="font-weight:600;font-size:0.88rem;margin-bottom:4px;">{mod['name']}</div>
          <div style="width:100%;height:8px;background:rgba(255,255,255,0.06);border-radius:4px;overflow:hidden;">
            <div style="width:{pct:.0f}%;height:100%;background:linear-gradient(90deg,{color},{color}bb);border-radius:4px;"></div>
          </div>
          <div style="font-size:0.7rem;color:#64748b;margin-top:4px;">{status}</div>
        </div>
      </div>
    </div>'''

# ============================================================
# 规则展示 / 编辑面板
# ============================================================
RULE_GROUPS = [
    ("📄 页面设置", "page", [
        ("纸张大小", "size", "text"),
        ("上边距 (cm)", "margin_top_cm", "number"),
        ("下边距 (cm)", "margin_bottom_cm", "number"),
        ("左边距 (cm)", "margin_left_cm", "number"),
        ("右边距 (cm)", "margin_right_cm", "number"),
    ]),
    ("📝 正文格式", "body", [
        ("中文字体", "cn_font", "font"),
        ("英文字体", "en_font", "font"),
        ("字号", "font_size", "size"),
        ("行距 (倍)", "line_spacing", "number"),
        ("首行缩进 (字符)", "first_indent_char", "number"),
    ]),
    ("📑 一级标题", "headings.h1", [
        ("字体", "font", "font"),
        ("字号", "size", "size"),
        ("加粗", "bold", "bool"),
    ]),
    ("📑 二级标题", "headings.h2", [
        ("字体", "font", "font"),
        ("字号", "size", "size"),
        ("加粗", "bold", "bool"),
    ]),
    ("📑 三级标题", "headings.h3", [
        ("字体", "font", "font"),
        ("字号", "size", "size"),
        ("加粗", "bold", "bool"),
    ]),
    ("📋 摘要要求", "abstract", [
        ("中文标题字体", "title_cn_font", "font"),
        ("中文标题字号", "title_cn_size", "size"),
        ("英文标题字体", "title_en_font", "font"),
        ("英文标题字号", "title_en_size", "size"),
        ("字数下限", "word_count_min", "int"),
        ("字数上限", "word_count_max", "int"),
        ("关键词最少", "keyword_count_min", "int"),
        ("关键词最多", "keyword_count_max", "int"),
    ]),
    ("📊 图表题注", "caption", [
        ("字体", "font", "font"),
        ("字号", "size", "size"),
        ("中英双语", "bilingual", "bool"),
    ]),
    ("📌 页眉", "header", [
        ("奇数页内容", "odd_page", "text"),
        ("偶数页内容", "even_page", "text"),
        ("奇偶页不同", "odd_even_different", "bool"),
    ]),
    ("📖 参考文献", "references", [
        ("最少篇数", "min_count", "int"),
        ("外文比例", "foreign_ratio", "number"),
        ("中文在前", "cn_before_en", "bool"),
    ]),
]

FONT_OPTIONS = ["宋体", "黑体", "楷体", "Times New Roman", "Arial"]
SIZE_OPTIONS = [
    "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
    "四号", "小四", "五号", "小五", "六号",
]


def _get_nested(d, dotted_key):
    """从 dict 中按 'headings.h1' 格式取值"""
    keys = dotted_key.split('.')
    for k in keys:
        d = d.get(k, {})
    return d


def _set_nested(d, dotted_key, field_key, value):
    """向 dict 中按 'headings.h1' 格式写值"""
    keys = dotted_key.split('.')
    ref = d
    for k in keys:
        ref = ref.setdefault(k, {})
    ref[field_key] = value


def _render_rules_panel(rules, editable=False):
    """渲染规则面板
    editable=False: 只读展示（基础版）
    editable=True: 可编辑（专业版/定制版）
    返回: 编辑后的 rules dict（editable=True 时）
    """
    edited = copy.deepcopy(rules) if editable else None

    for group_label, group_key, fields in RULE_GROUPS:
        with st.expander(group_label, expanded=False):
            group_data = _get_nested(rules, group_key)
            cols = st.columns(2)
            for i, (label, field_key, field_type) in enumerate(fields):
                val = group_data.get(field_key, '')
                with cols[i % 2]:
                    if not editable:
                        st.text_input(label, value=str(val), disabled=True,
                                      key=f"rule_{group_key}_{field_key}")
                    else:
                        if field_type == "font":
                            new_val = st.selectbox(
                                label, FONT_OPTIONS,
                                index=FONT_OPTIONS.index(val) if val in FONT_OPTIONS else 0,
                                key=f"edit_{group_key}_{field_key}")
                        elif field_type == "size":
                            new_val = st.selectbox(
                                label, SIZE_OPTIONS,
                                index=SIZE_OPTIONS.index(val) if val in SIZE_OPTIONS else 0,
                                key=f"edit_{group_key}_{field_key}")
                        elif field_type == "bool":
                            new_val = st.checkbox(label, value=bool(val),
                                                  key=f"edit_{group_key}_{field_key}")
                        elif field_type == "int":
                            new_val = st.number_input(
                                label, value=int(val) if val else 0,
                                step=1, key=f"edit_{group_key}_{field_key}")
                        elif field_type == "number":
                            new_val = st.number_input(
                                label, value=float(val) if val else 0.0,
                                step=0.1, format="%.2f",
                                key=f"edit_{group_key}_{field_key}")
                        else:
                            new_val = st.text_input(label, value=str(val),
                                                    key=f"edit_{group_key}_{field_key}")
                        if edited is not None:
                            _set_nested(edited, group_key, field_key, new_val)

    return edited


# ============================================================
# 侧边栏（精简，仅管理员）
# ============================================================
# 管理员后台（放在页面最底部，折叠隐藏，普通用户看不到）
def _render_admin_panel():
    """渲染管理员面板，放在页面最底部"""
    with st.expander("管理", expanded=False):
        admin_pwd = st.text_input("密码", type="password", key="admin_pwd")
        if admin_pwd == ADMIN_PWD:
            st.success("已登录")
            col_a, col_b = st.columns(2)
            with col_a:
                gen_n = st.number_input("生成数量", min_value=1, max_value=100, value=10)
            with col_b:
                gen_tier = st.selectbox("套餐", ['lite', 'basic', 'pro', 'fix'])
            if st.button("生成兑换码", use_container_width=True):
                new_codes = generate_codes(gen_n, gen_tier)
                st.code('\n'.join(new_codes))
            codes = load_codes_safe()
            unused = sum(1 for c in codes.values() if not c['used'])
            used = sum(1 for c in codes.values() if c['used'])
            c1, c2 = st.columns(2)
            c1.metric("可用", unused)
            c2.metric("已用", used)
            if st.button("查看全部", use_container_width=True):
                for code, info in codes.items():
                    s = "已用" if info['used'] else "可用"
                    st.text(f"{code} [{s}] {info.get('tier','basic')}")

# ============================================================
# ---- 主页面 ----
# ============================================================

# 倒计时 banner（基于答辩典型日期 5/31 估算，用户可忽略）
_today = datetime.now()
_defense_dates = {'研究生': datetime(_today.year, 5, 31), '本科': datetime(_today.year, 6, 15)}
_default_defense = _defense_dates[st.session_state.get('edu_level', '研究生')]
if _default_defense < _today:
    _default_defense = _default_defense.replace(year=_today.year + 1)
_days_to_defense = (_default_defense - _today).days
if 0 < _days_to_defense <= 90:
    st.markdown(f'''
    <div style="background:linear-gradient(90deg,rgba(239,68,68,0.12),rgba(234,179,8,0.12));
        border:1px solid rgba(239,68,68,0.3);border-radius:10px;
        padding:10px 18px;margin:12px 0;text-align:center;font-size:0.88rem;
        color:var(--text-primary);">
        ⏳ 距离典型毕业答辩日期（{_default_defense.strftime('%m月%d日')}）还剩
        <b style="color:#ef4444;font-size:1.1rem;">{_days_to_defense}</b> 天
        &nbsp;·&nbsp; 平均修改需 4-8 小时，现在查比答辩前一晚强
    </div>
    ''', unsafe_allow_html=True)

# Hero
st.markdown('''
<div class="hero">
    <h1>论文格式一键体检</h1>
    <p>本科 / 研究生毕业论文通用 · 答辩前最后一道关 · 5分钟出体检报告</p>
    <div class="highlights">
        <div class="hl-item"><div class="hl-num" style="color:#3b82f6;">30+</div><div class="hl-label">答辩高频扣分雷区</div></div>
        <div class="hl-item"><div class="hl-num" style="color:#3b82f6;">14大</div><div class="hl-label">模块全覆盖</div></div>
        <div class="hl-item"><div class="hl-num" style="color:#3b82f6;">5分钟</div><div class="hl-label">出报告</div></div>
        <div class="hl-item"><div class="hl-num" style="color:#3b82f6;">Beta版</div><div class="hl-label">不吹"全查"</div></div>
    </div>
</div>
''', unsafe_allow_html=True)

# 学历切换 + 上传区
st.markdown('<div class="upload-zone">', unsafe_allow_html=True)

# 学历档位（影响参考文献数量、摘要字数等规则）
col_edu1, col_edu2 = st.columns([1, 4])
with col_edu1:
    st.markdown('<div style="padding-top:8px;font-size:0.9rem;color:#94a3b8;">学历档位：</div>',
        unsafe_allow_html=True)
with col_edu2:
    edu_level = st.radio("学历档位", ['研究生', '本科'],
        index=0, horizontal=True, label_visibility="collapsed",
        help="本科论文和研究生论文在参考文献数量、摘要字数等要求上不同")
st.session_state['edu_level'] = edu_level

col_up, col_title = st.columns([3, 2])
with col_up:
    uploaded_file = st.file_uploader("上传论文 (.docx / .doc)", type=['docx', 'doc'],
        help="支持 .docx 和 .doc 格式，最大 200MB", label_visibility="collapsed")
with col_title:
    thesis_title = st.text_input("论文题目（可选，用于页眉校验）",
        placeholder="如：基于深度学习的小麦病害图像识别研究",
        label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

# 3 条信任钩子（替换掉之前的"2400+"自吹数字）
st.markdown('''
<div style="display:flex;flex-wrap:wrap;justify-content:center;gap:10px;margin:12px 0 20px;font-size:0.82rem;">
    <div style="padding:6px 12px;background:rgba(59,130,246,0.08);border:1px solid rgba(59,130,246,0.25);border-radius:20px;color:var(--text-secondary);">
        🔒 文件仅在内存中解析，24 小时自动删除，不用于训练 AI
    </div>
    <div style="padding:6px 12px;background:rgba(34,197,94,0.08);border:1px solid rgba(34,197,94,0.25);border-radius:20px;color:var(--text-secondary);">
        ✓ 覆盖毕业论文常见 30+ 格式雷区（Beta 版，不保证查出全部）
    </div>
    <div style="padding:6px 12px;background:rgba(234,179,8,0.08);border:1px solid rgba(234,179,8,0.25);border-radius:20px;color:var(--text-secondary);">
        💸 免费预览 3 条问题 · 付费后报告不对无理由退款
    </div>
</div>
''', unsafe_allow_html=True)

# ---- 定制版：格式规范上传入口（策略B：所有人可见，解析时拦截）----
with st.expander("📎 上传学校格式规范（定制版专属）", expanded=False):
    st.caption("支持 PDF、图片，AI 自动识别格式要求")
    template_file = st.file_uploader(
        "选择格式规范文件",
        type=['pdf', 'png', 'jpg', 'jpeg'],
        key="template_upload")
    if template_file:
        if not _get_tier_config()['ai_parse']:
            st.warning("🔒 AI 解析格式规范为定制版专属功能，请购买定制版套餐后使用")
        else:
            with st.spinner("AI 正在解析格式规范..."):
                from template_parser import parse_template
                try:
                    parsed_rules = parse_template(template_file.getvalue(), template_file.name)
                    st.session_state['custom_rules'] = parsed_rules
                    st.success(f"✅ 已识别 {parsed_rules.get('school_name', '未知')} 的格式规则")
                except Exception as e:
                    st.error(f"解析失败: {e}")

# ============================================================
# 免费次数限制（浏览器指纹 + session_state）
# ============================================================
FREE_CHECK_LIMIT = 2  # 免费版最多检查 2 次

def _get_usage_count():
    """获取免费使用次数（优先从 localStorage 同步）"""
    _sync_usage_from_local_storage()
    return st.session_state.get('free_usage', 0)

def _increment_usage():
    new_count = st.session_state.get('free_usage', 0) + 1
    st.session_state['free_usage'] = new_count
    _persist_usage_to_local_storage(new_count)

# 跨刷新持久计数：用 streamlit-js-eval 读写 localStorage
from streamlit_js_eval import streamlit_js_eval

def _sync_usage_from_local_storage():
    """从 localStorage 读取免费使用次数，同步到 session_state"""
    if 'usage_synced' not in st.session_state:
        stored = streamlit_js_eval(
            js_expressions="parseInt(localStorage.getItem('fmt_free_count') || '0')",
            key="read_usage")
        if stored is not None:
            st.session_state['free_usage'] = int(stored)
        st.session_state['usage_synced'] = True

def _persist_usage_to_local_storage(count):
    """将使用次数写入 localStorage"""
    streamlit_js_eval(
        js_expressions=f"localStorage.setItem('fmt_free_count', '{count}')",
        key=f"write_usage_{count}")

# ============================================================
# 检查流程
# ============================================================
_can_check = False
_skip_increment = False  # 免费次数已用完时跳过计数递增，但仍允许检查（付费墙会拦截）
if uploaded_file is not None:
    usage = _get_usage_count()
    unlocked_session = st.session_state.get('unlocked', False)
    in_payment_flow = 'pay_tier' in st.session_state or 'auto_code' in st.session_state
    if usage >= FREE_CHECK_LIMIT and not unlocked_session and not in_payment_flow:
        # 免费次数用完：仍然运行检查，让用户看到预览和付费墙，但不再递增计数
        _can_check = True
        _skip_increment = True
    elif unlocked_session:
        # 付费用户：检查复查次数是否超限
        tc = _get_tier_config()
        recheck_count = st.session_state.get('recheck_count', 0)
        recheck_limit = tc['recheck_limit']
        if recheck_limit != -1 and recheck_count > recheck_limit:
            st.warning(f"已用完 {recheck_limit} 次复查机会，升级套餐可获得更多复查次数")
        else:
            _can_check = True
    else:
        _can_check = True

if uploaded_file is not None and _can_check:
    # 用文件内容的 hash 判断是否同一份论文，避免 rerun 时重复检查
    _file_bytes = uploaded_file.getvalue()
    _file_hash = hashlib.md5(_file_bytes).hexdigest()
    _custom_rules = st.session_state.get('custom_rules')
    _rules_hash = hashlib.md5(json.dumps(_custom_rules, sort_keys=True, default=str).encode()).hexdigest() if _custom_rules else 'default'
    _cache = st.session_state.get('check_cache', {})
    _cache_hit = (_cache.get('file_hash') == _file_hash
                  and _cache.get('thesis_title') == (thesis_title or None)
                  and _cache.get('rules_hash') == _rules_hash)

    if _cache_hit:
        # 缓存命中：直接读取上次结果，不重新检查
        data = _cache['data']
        html_content = _cache['html_content']
        report_id = _cache['report_id']
    else:
        # 首次检查或文件变化：执行检查并缓存
        if not _skip_increment and not st.session_state.get('unlocked', False) and 'pay_tier' not in st.session_state and 'auto_code' not in st.session_state:
            # 免费用户：递增免费计数
            _increment_usage()
        elif st.session_state.get('unlocked', False):
            # 付费用户：递增复查计数（首次检查不算复查）
            if st.session_state.get('check_cache'):
                st.session_state['recheck_count'] = st.session_state.get('recheck_count', 0) + 1

        _is_doc = uploaded_file.name.lower().endswith('.doc') and not uploaded_file.name.lower().endswith('.docx')
        _suffix = '.doc' if _is_doc else '.docx'
        with tempfile.NamedTemporaryFile(delete=False, suffix=_suffix) as tmp:
            tmp.write(_file_bytes)
            tmp_path = tmp.name

        if _is_doc:
            st.info("检测到 .doc 格式，正在转换为 .docx ...")
            out_dir = os.path.dirname(tmp_path)
            result = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'docx', '--outdir', out_dir, tmp_path],
                capture_output=True, timeout=120)
            converted_path = tmp_path.rsplit('.', 1)[0] + '.docx'
            if result.returncode != 0 or not os.path.exists(converted_path):
                st.error("❌ .doc 文件转换失败，请用 Word 打开后另存为 .docx 格式再上传")
                os.unlink(tmp_path)
                st.stop()
            os.unlink(tmp_path)
            tmp_path = converted_path

        # 修复版所需的原文件字节会在下面随 check_cache 一起落盘，这里不再暴露临时路径
        html_path = None
        try:
            progress_bar = st.progress(0, text="正在解析论文文档...")
            status_text = st.empty()
            t0 = time.time()
            # 规则优先级：AI 解析的 custom_rules > 学历档位默认规则 > 硕士默认规则
            _runtime_rules = st.session_state.get('custom_rules')
            if not _runtime_rules:
                _runtime_rules = get_default_rules(st.session_state.get('edu_level', '研究生') and
                                                   ('本科' if st.session_state.get('edu_level') == '本科' else '硕士'))
            checker = ThesisChecker(tmp_path, thesis_title=thesis_title or None,
                                   rules=_runtime_rules)
            progress_bar.progress(5, text="文档解析完成，开始格式审查...")

            def _on_progress(step, total, name):
                pct = int(5 + 85 * step / total)
                if step < total:
                    progress_bar.progress(pct, text=f"正在检查：{name}（{step+1}/{total}）")
                else:
                    progress_bar.progress(95, text="正在生成报告...")

            checker.run_all_checks(progress_callback=_on_progress)
            elapsed = time.time() - t0
            data = checker.get_report_data()
            html_path = tmp_path.replace('.docx', '_report.html')
            checker.generate_html_report(html_path)
            with open(html_path, 'r', encoding='utf-8') as f:
                html_content = f.read()
            progress_bar.progress(100, text=f"审查完成！耗时 {elapsed:.1f} 秒")
            time.sleep(0.5)
            progress_bar.empty()
            status_text.empty()
        finally:
            try:
                os.unlink(tmp_path)
                if html_path and os.path.exists(html_path):
                    os.unlink(html_path)
            except Exception:
                pass

        report_id = f"FMT-{datetime.now().strftime('%Y%m%d')}-{hashlib.md5(_file_bytes[:1024]).hexdigest()[:6].upper()}"

        # 写入缓存（带原文件字节供修复版使用，避免临时文件被 finally 清理后丢失）
        st.session_state['check_cache'] = {
            'file_hash': _file_hash,
            'thesis_title': thesis_title or None,
            'rules_hash': _rules_hash,
            'data': data,
            'html_content': html_content,
            'report_id': report_id,
            'file_bytes': _file_bytes,
        }

    st.success(f"审查完成！报告编号 {report_id}")

    # ========== 方案 B：免费藏分 / 付费揭晓 ==========
    _unlocked_now = st.session_state.get('unlocked', False)
    _err_cnt = data['error_count']
    _warn_cnt = data['warning_count']

    # ---- 致命问题 Hero（首屏核心，损失厌恶 framing）----
    if not _unlocked_now:
        if _err_cnt > 0:
            st.markdown(f'''
            <div style="background:linear-gradient(135deg,rgba(239,68,68,0.15),rgba(239,68,68,0.05));
                border:1px solid rgba(239,68,68,0.45);border-radius:14px;
                padding:20px 24px;margin:16px 0 20px;text-align:center;">
                <div style="font-size:1.1rem;color:#fca5a5;margin-bottom:4px;">⚠ 检测到</div>
                <div style="font-size:3.2rem;font-weight:900;color:#ef4444;line-height:1;margin:6px 0;">
                    {_err_cnt}
                </div>
                <div style="font-size:1.1rem;font-weight:700;color:#f87171;margin-bottom:8px;">
                    处致命格式问题
                </div>
                <div style="font-size:0.88rem;color:var(--text-secondary);line-height:1.6;">
                    这些问题可能导致 <b style="color:#fca5a5;">答辩现场被要求返修</b> 或 <b style="color:#fca5a5;">学校抽检被打回</b><br>
                    延毕半年 ≈ 房租 3000×6 + 机会成本，<b>24.9 元相当于格式保险</b>
                </div>
            </div>
            ''', unsafe_allow_html=True)
        else:
            st.markdown(f'''
            <div style="background:linear-gradient(135deg,rgba(234,179,8,0.12),rgba(234,179,8,0.04));
                border:1px solid rgba(234,179,8,0.35);border-radius:14px;
                padding:20px 24px;margin:16px 0 20px;text-align:center;">
                <div style="font-size:1rem;color:#fde68a;margin-bottom:4px;">✓ 未检出致命错误</div>
                <div style="font-size:2rem;font-weight:800;color:#eab308;">
                    但还有 {_warn_cnt} 处格式警告
                </div>
                <div style="font-size:0.85rem;color:var(--text-secondary);margin-top:6px;">
                    这些问题不一定被打回，但会被导师/评审老师圈红
                </div>
            </div>
            ''', unsafe_allow_html=True)
        # 分数环打码：只显示问号 + 解锁后可见提示
        st.markdown('''
        <div style="display:flex;gap:24px;align-items:center;justify-content:center;flex-wrap:wrap;margin:8px 0 4px;">
            <div style="width:140px;height:140px;border:8px solid rgba(148,163,184,0.25);
                border-radius:50%;display:flex;align-items:center;justify-content:center;
                background:rgba(148,163,184,0.06);">
                <div style="text-align:center;">
                    <div style="font-size:2.2rem;font-weight:900;color:#64748b;">?</div>
                    <div style="font-size:0.68rem;color:#94a3b8;margin-top:-4px;">总分待解锁</div>
                </div>
            </div>
            <div style="font-size:0.85rem;color:var(--text-secondary);line-height:1.9;max-width:320px;">
                🎯 <b>解锁任意套餐，查看</b>：<br>
                &nbsp;&nbsp;• 总分 + 评级（A/B/C/D）<br>
                &nbsp;&nbsp;• 你的论文在同届中的排名百分位<br>
                &nbsp;&nbsp;• 修复后的预估分数提升<br>
                &nbsp;&nbsp;• 每处问题的精确位置 + 修改建议
            </div>
        </div>
        ''', unsafe_allow_html=True)
    else:
        # 付费用户：揭晓分数 + 百分位 + 修复动力
        col_ring, col_metrics = st.columns([1, 2])
        with col_ring:
            st.markdown(render_score_ring(data['total_score'], data['max_score'], data['grade']),
                unsafe_allow_html=True)
            pct_score = data['pct']
            if pct_score >= 90: beat_pct = 85
            elif pct_score >= 80: beat_pct = 65
            elif pct_score >= 70: beat_pct = 40
            elif pct_score >= 60: beat_pct = 20
            elif pct_score >= 50: beat_pct = 10
            else: beat_pct = 5
            st.markdown(f'<p style="text-align:center;font-size:0.85rem;color:#94a3b8;margin-top:4px;">'
                f'你的论文格式超过了 <b style="color:#818cf8;">{beat_pct}%</b> 的论文</p>',
                unsafe_allow_html=True)
            # 修复动力：距 A+ 还差多少（未到 A 的显示距 A 的距离）
            if data['total_score'] < 97 and _err_cnt > 0:
                target = 'A+' if pct_score >= 90 else ('A' if pct_score >= 80 else '下一档')
                st.markdown(f'<p style="text-align:center;font-size:0.82rem;color:#fbbf24;margin-top:4px;">'
                    f'距离 <b>{target}</b> 还差修复 {_err_cnt} 处致命错误</p>',
                    unsafe_allow_html=True)
            if _err_cnt > 0:
                st.markdown(f'<p style="text-align:center;font-size:0.78rem;color:#f87171;margin-top:2px;">'
                    f'⚠ 建议提交前逐一修复</p>', unsafe_allow_html=True)

            # 修复后预估分（可修复问题按权重估算，给出修复版的升级钩子）
            _fixable = data.get('fixable_count', 0)
            if _fixable > 0 and data['total_score'] < data['max_score']:
                # 粗估：每个可修复的 error 加回 1.5 分，warning 加回 0.7 分，capped
                fixable_errors = sum(1 for i in issues
                    if i.get('fixable') and i.get('severity') == 'error')
                fixable_warnings = sum(1 for i in issues
                    if i.get('fixable') and i.get('severity') == 'warning')
                est_gain = min(fixable_errors * 1.5 + fixable_warnings * 0.7,
                              data['max_score'] - data['total_score'])
                est_score = round(data['total_score'] + est_gain, 1)
                est_grade = 'A+' if est_score >= 97 else ('A' if est_score >= 85 else 'B')
                _current_tier = st.session_state.get('user_tier', 'basic')
                cta_line = ('✨ 修复版可自动修复这些问题' if _current_tier != 'fix'
                            else '✓ 你已有修复版，点击下方"一键修复"')
                st.markdown(f'''
                <div style="background:linear-gradient(135deg,rgba(34,197,94,0.1),rgba(34,197,94,0.02));
                    border:1px solid rgba(34,197,94,0.35);border-radius:10px;
                    padding:12px 14px;margin-top:12px;text-align:center;">
                    <div style="font-size:0.78rem;color:#86efac;">修复后预估提升</div>
                    <div style="font-size:1.3rem;font-weight:800;color:#22c55e;margin:4px 0;">
                        {data["total_score"]} → {est_score} <span style="font-size:0.9rem;color:#86efac;">({est_grade})</span>
                    </div>
                    <div style="font-size:0.72rem;color:var(--text-muted);">
                        {_fixable} 项可自动修复 · {cta_line}
                    </div>
                </div>
                ''', unsafe_allow_html=True)
        with col_metrics:
            m1, m2, m3 = st.columns(3)
            m1.metric("严重错误", data['error_count'])
            m2.metric("格式警告", data['warning_count'])
            m3.metric("优化建议", data['info_count'])
            m4, m5, m6 = st.columns(3)
            m4.metric("段落数", data['total_paras'])
            m5.metric("表格数", data['total_tables'])
            m6.metric("图片数", data['total_images'])

    # 免费用户也显示问题数量指标（但不显示段落/表格/图片的中性指标避免平静感）
    if not _unlocked_now:
        m1, m2, m3 = st.columns(3)
        m1.metric("严重错误", _err_cnt, delta=None,
                  help="必须修复，否则答辩/抽检容易被打回")
        m2.metric("格式警告", _warn_cnt, help="会被导师/评审老师指出返修")
        m3.metric("优化建议", data['info_count'], help="非强制，但能让论文更规范")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ========== 模块评分卡片网格 ==========
    mods = data['modules']
    unlocked_mods = st.session_state.get('unlocked', False)

    if unlocked_mods:
        # 付费用户：显示完整模块分数
        st.markdown("#### 各模块得分")
        grid_html = '<div class="mod-grid">'
        for mod in mods:
            grid_html += render_module_card(mod)
        grid_html += '</div>'
        st.markdown(grid_html, unsafe_allow_html=True)
    else:
        # 免费用户：统计摘要 + 锁定卡片（不泄露具体哪个模块有问题）
        pass_count = sum(1 for m in mods if m['pct'] >= 90)
        warn_count = sum(1 for m in mods if 40 <= m['pct'] < 90)
        fail_count = sum(1 for m in mods if m['pct'] < 40)
        st.markdown("#### 14 个维度扫描完成")
        summary_html = f'''<div style="display:flex;gap:16px;margin-bottom:16px;flex-wrap:wrap;">
            <div class="glass-card" style="flex:1;min-width:120px;padding:16px;text-align:center;">
                <div style="font-size:1.8rem;font-weight:800;color:#10b981;">{pass_count}</div>
                <div style="font-size:0.8rem;color:var(--text-secondary);">维度通过</div>
            </div>
            <div class="glass-card" style="flex:1;min-width:120px;padding:16px;text-align:center;">
                <div style="font-size:1.8rem;font-weight:800;color:#f59e0b;">{warn_count}</div>
                <div style="font-size:0.8rem;color:var(--text-secondary);">需要改进</div>
            </div>
            <div class="glass-card" style="flex:1;min-width:120px;padding:16px;text-align:center;">
                <div style="font-size:1.8rem;font-weight:800;color:#ef4444;">{fail_count}</div>
                <div style="font-size:0.8rem;color:var(--text-secondary);">严重不足</div>
            </div>
        </div>'''
        st.markdown(summary_html, unsafe_allow_html=True)

        # 锁定的模块卡片网格
        grid_html = '<div class="mod-grid">'
        for mod in mods:
            grid_html += render_module_card(mod, locked=True)
        grid_html += '</div>'
        st.markdown(grid_html, unsafe_allow_html=True)
        st.markdown('<p style="text-align:center;font-size:0.82rem;color:#818cf8;margin-top:8px;">'
            '🔒 解锁任意套餐，查看具体哪些维度不达标</p>', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ========== 问题列表 ==========
    issues = data['issues']
    st.markdown(f"#### 问题详情（共 {len(issues)} 条）")

    # 免费预览：冰山一角策略 — 固定优先 2 条严重错误 + 1 条警告，制造 Zeigarnik 效应
    # （未完成感 > 随机抽样，让用户看到"真问题"的样子后对剩余内容产生强烈好奇）
    FREE_LIMIT = 3
    errors_only = [i for i in issues if i['severity'] == 'error']
    warnings_only = [i for i in issues if i['severity'] == 'warning']
    infos_only = [i for i in issues if i['severity'] == 'info']

    free_preview = []
    # 先塞 2 条严重错误（按模块优先级：编号>标题>图表>正文）
    priority_modules = ['编号规范', '标题层级', '图表规范', '正文格式', '参考文献', '封面', '目录']
    errors_sorted = sorted(errors_only,
        key=lambda i: (priority_modules.index(i['module']) if i['module'] in priority_modules else 99,
                       i.get('para_index', 0)))
    free_preview.extend(errors_sorted[:2])
    # 再塞 1 条警告
    if len(free_preview) < FREE_LIMIT and warnings_only:
        free_preview.append(warnings_only[0])
    # 如果严重错误不够 2 条，用警告补
    while len(free_preview) < FREE_LIMIT:
        candidates = [i for i in warnings_only + infos_only if i not in free_preview]
        if not candidates:
            break
        free_preview.append(candidates[0])

    for issue in free_preview:
        st.markdown(render_issue(issue), unsafe_allow_html=True)

    # ========== 付费墙 / 完整报告 ==========
    if len(issues) > FREE_LIMIT:
        unlocked = st.session_state.get('unlocked', False)

        if not unlocked:
            # 付费墙遮罩（损失厌恶 framing，锚定答辩成本）
            # 计算还有多少处未展示的严重错误（总错 - 已展示的）
            shown_errors = sum(1 for i in free_preview if i['severity'] == 'error')
            shown_warnings = sum(1 for i in free_preview if i['severity'] == 'warning')
            hidden_errors = _err_cnt - shown_errors
            hidden_warnings = _warn_cnt - shown_warnings
            hidden_total = len(issues) - len(free_preview)

            # 核心恐惧文案：优先提严重错误，没有严重错误时提警告
            if hidden_errors > 0:
                fear_line = f'<b style="color:#ef4444;">还有 {hidden_errors} 处致命错误</b>未展示 — 其中任何一处都可能导致答辩返修'
            elif hidden_warnings > 0:
                fear_line = f'<b style="color:#fbbf24;">还有 {hidden_warnings} 处格式警告</b>未展示 — 这些会被导师/评审老师圈红要求返工'
            else:
                fear_line = f'还有 {hidden_total} 条问题未展示'

            st.markdown(f'''
            <div class="paywall" style="text-align:center;padding:24px 20px;">
                <div style="font-size:2.2rem;margin-bottom:4px;">🔒</div>
                <div style="font-size:1.15rem;font-weight:700;color:var(--text-primary);margin-bottom:8px;line-height:1.6;">
                    {fear_line}
                </div>
                <div style="color:var(--text-secondary);font-size:0.92rem;margin-bottom:12px;line-height:1.7;">
                    答辩被打回 = 返修几周 + 延期答辩 + 重新约导师<br>
                    <b style="color:#fde047;">24.9 元 ≈ 2 天外卖</b>，但省下的可能是几个月时间
                </div>
                <div style="color:var(--text-muted);font-size:0.82rem;margin-bottom:12px;padding:8px 12px;background:rgba(148,163,184,0.08);border-radius:8px;display:inline-block;">
                    ⏱ 5 分钟付费 → 全部问题定位到页码行号 → 今晚就能改完
                </div>
            </div>''', unsafe_allow_html=True)

            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

            # ---- 套餐选择（三档）----
            st.markdown("#### 毕业季特惠 · 选择套餐")
            t1, t2, t3 = st.columns(3)
            with t1:
                st.markdown('''<div class="glass-card tier-free" style="text-align:center;padding:24px 16px;">
                    <div style="font-size:1rem;font-weight:700;">极简版</div>
                    <div style="margin:10px 0;">
                        <span class="original-price">原价 19.9 元</span><br>
                        <span class="price" style="font-size:2rem;font-weight:800;color:#94a3b8;">9.9 元</span>
                        <span class="discount-badge">毕业季半价</span>
                    </div>
                    <div style="font-size:0.78rem;color:#94a3b8;line-height:1.9;text-align:left;padding:0 8px;">
                        <i style="color:#cbd5e1;">要一份能看的清单，自己改</i><br>
                        14 个模块 + 30+ 格式雷区扫描<br>
                        查看前 5 条问题详情<br>
                        含问题位置定位 · 1 次检查</div>
                </div>''', unsafe_allow_html=True)
                pick_lite = st.button("选择极简版", key="pick_lite", use_container_width=True)
            with t2:
                st.markdown('''<div class="glass-card tier-basic" style="text-align:center;padding:24px 16px;">
                    <div style="font-size:1rem;font-weight:700;">基础版</div>
                    <div style="margin:10px 0;">
                        <span class="original-price">原价 49.9 元</span><br>
                        <span class="price" style="font-size:2rem;font-weight:800;">24.9 元</span>
                        <span class="discount-badge">毕业季5折</span>
                    </div>
                    <div style="font-size:0.78rem;color:#94a3b8;line-height:1.9;text-align:left;padding:0 8px;">
                        <i style="color:#cbd5e1;">清单 + 每条问题修改示范</i><br>
                        极简版全部功能<br>
                        <b>查看全部问题详情</b> · 筛选<br>
                        下载完整 HTML 报告 · 3 次复查</div>
                </div>''', unsafe_allow_html=True)
                pick_basic = st.button("选择基础版", key="pick_basic", use_container_width=True)
            with t3:
                st.markdown('''<div class="glass-card tier-pro" style="text-align:center;padding:24px 16px;">
                    <div style="font-size:1rem;font-weight:700;">专业版 <span class="recommend-badge">推荐</span></div>
                    <div style="margin:10px 0;">
                        <span class="original-price">原价 99.9 元</span><br>
                        <span class="price" style="font-size:2rem;font-weight:800;">49.9 元</span>
                        <span class="discount-badge">毕业季5折</span>
                    </div>
                    <div style="font-size:0.78rem;color:#94a3b8;line-height:1.9;text-align:left;padding:0 8px;">
                        <i style="color:#cbd5e1;">定位到页码行号，适合反复改 3 版以上</i><br>
                        基础版全部功能<br>
                        <b>每条问题附修改建议</b><br>
                        <b>自定义编辑全部检查规则</b> · 适配任意学校<br>
                        不限次复查 · 优先客服</div>
                </div>''', unsafe_allow_html=True)
                pick_pro = st.button("选择专业版", key="pick_pro", type="primary", use_container_width=True)

            # ---- 修复版卡片 ----
            st.markdown('''
            <div style="max-width:360px;margin:16px auto;">
                <div class="glass-card tier-pro" style="text-align:center;padding:24px 16px;border-color:rgba(234,179,8,0.6);">
                    <div style="font-size:1rem;font-weight:700;">修复版 <span class="recommend-badge" style="background:linear-gradient(135deg,#f59e0b,#ef4444);">省300元</span></div>
                    <div style="margin:10px 0;">
                        <span class="original-price">原价 199.9 元</span><br>
                        <span class="price" style="font-size:2rem;font-weight:800;">99.9 元</span>
                        <span class="discount-badge">毕业季5折</span>
                    </div>
                    <div style="font-size:0.78rem;color:#94a3b8;line-height:1.9;text-align:left;padding:0 8px;">
                        <i style="color:#cbd5e1;">直接给你改好的 docx · 省 300-500 元人工</i><br>
                        专业版全部功能<br>
                        <b>一键自动修复格式问题</b><br>
                        <b>下载修复后的论文文件</b><br>
                        修复前预览 · 修复后复查</div>
                    <div style="font-size:0.72rem;color:#f59e0b;margin-top:8px;">人工改格式 300-500 元 · 本工具 99.9 元 · 周末解放</div>
                </div>
            </div>''', unsafe_allow_html=True)
            pick_fix = st.button("选择修复版", key="pick_fix", use_container_width=True)

            st.caption("邀请同学使用你的专属邀请码购买，双方各返 5 元")

            # 选定套餐后弹出付款区（用按钮当前帧判断，不持久化到 session_state）
            just_picked = None
            if pick_lite: just_picked = ("极简版", "9.9", "lite")
            elif pick_basic: just_picked = ("基础版", "24.9", "basic")
            elif pick_pro: just_picked = ("专业版", "49.9", "pro")
            elif pick_fix: just_picked = ("修复版", "99.9", "fix")

            if just_picked:
                # 切换套餐时清除之前的兑换码，防止套餐和码不匹配
                old_tier = st.session_state.get('pay_tier')
                if old_tier and old_tier != just_picked[0]:
                    st.session_state.pop('auto_code', None)
                st.session_state['pay_tier'] = just_picked[0]
                st.session_state['pay_price'] = just_picked[1]

            # 只在用户选了套餐后显示付款区
            if 'pay_tier' in st.session_state:
                tier_name = st.session_state['pay_tier']
                tier_price = st.session_state['pay_price']

                st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
                st.markdown(f"#### 付款 · {tier_name}（{tier_price}元）")

                col_qr, col_unlock = st.columns([1, 1], gap="large")
                with col_qr:
                    # 多路径查找二维码图片
                    pay_img = None
                    for candidate in [
                        os.path.join(os.path.dirname(__file__), 'zhifubao.jpg'),
                        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'zhifubao.jpg'),
                        'zhifubao.jpg',
                    ]:
                        if os.path.exists(candidate):
                            pay_img = candidate
                            break

                    if pay_img:
                        st.image(pay_img, width=220, caption=f"支付宝扫码 · {tier_price}元")
                    else:
                        st.warning(f"⚠️ 二维码加载失败，请添加微信 **l8811925** 转账 {tier_price} 元")

                    # 生成付款Token（基于报告编号+套餐，不含解锁能力）
                    if st.button("我已付款", key="paid_btn", use_container_width=True):
                        token_raw = f"{report_id}-{tier_name}-{_get_session_id()}"
                        token = "PAY-" + hashlib.md5(token_raw.encode()).hexdigest()[:8].upper()
                        st.session_state['pay_token'] = token

                    if 'pay_token' in st.session_state:
                        token = st.session_state['pay_token']
                        st.markdown(f"""
                        <div style="background:rgba(59,130,246,0.1);border:1px solid rgba(59,130,246,0.3);
                            border-radius:10px;padding:16px;margin-top:8px;text-align:center;">
                            <div style="font-size:0.85rem;color:var(--text-secondary);margin-bottom:8px;">
                                你的付款凭证 Token</div>
                            <div style="font-size:1.3rem;font-weight:800;color:var(--accent-blue);
                                letter-spacing:2px;margin-bottom:12px;">{token}</div>
                            <div style="font-size:0.88rem;color:var(--text-primary);line-height:1.8;">
                                添加微信 <b style="color:var(--accent-blue);">l8811925</b><br>
                                发送 <b>付款截图</b> + Token <b style="color:var(--accent-blue);">{token}</b><br>
                                确认后发你兑换码，输入右侧即可解锁</div>
                            <div style="font-size:0.78rem;color:var(--text-muted);margin-top:8px;">
                                工作时间 5 分钟内回复</div>
                        </div>""", unsafe_allow_html=True)
                    else:
                        st.caption("付完款点击上方按钮，获取付款凭证")

                with col_unlock:
                    st.markdown("##### 输入兑换码解锁")
                    st.markdown("""
                    <div style="font-size:0.85rem;color:var(--text-secondary);margin-bottom:12px;line-height:1.6;">
                        收到兑换码后，粘贴到下方输入框即可解锁完整报告
                    </div>""", unsafe_allow_html=True)

                    code_input = st.text_input("兑换码", placeholder="FMT-XXXX-XXXX",
                        label_visibility="collapsed")
                    if st.button("解锁完整报告", type="primary", use_container_width=True):
                        if code_input:
                            ok, msg = verify_code(code_input,
                                report_id=report_id, filename=uploaded_file.name)
                            if ok:
                                st.session_state['unlocked'] = True
                                # 从兑换码中读取套餐类型，写入 session
                                codes = load_codes()
                                code_upper = code_input.strip().upper()
                                if code_upper in codes:
                                    st.session_state['user_tier'] = codes[code_upper].get('tier', 'basic')
                                st.session_state.pop('pay_tier', None)
                                st.session_state.pop('pay_price', None)
                                st.session_state.pop('auto_code', None)
                                st.rerun()
                            else:
                                st.error(msg)
                        else:
                            st.warning("请输入兑换码")
                    st.markdown("""
                    <div style="font-size:0.8rem;color:var(--text-muted);margin-top:12px;line-height:1.8;">
                        解锁后包含（按套餐分级）：<br>
                        &nbsp;&nbsp;极简版：前 5 条问题详情<br>
                        &nbsp;&nbsp;基础版：全部问题 + 筛选 + 下载报告<br>
                        &nbsp;&nbsp;专业版：全部 + 修改建议 + 自定义规则
                    </div>""", unsafe_allow_html=True)

        else:
            # ========== 已解锁报告（根据套餐权限分级展示）==========
            tc = _get_tier_config()
            tier_name = st.session_state.get('user_tier', 'basic')
            tier_labels = {'lite': '极简版', 'basic': '基础版', 'pro': '专业版', 'fix': '修复版'}
            st.markdown(f"**已解锁 · {tier_labels.get(tier_name, '基础版')}**")

            # ---- 规则面板（根据套餐权限展示）----
            current_rules = st.session_state.get('custom_rules') or DEFAULT_RULES
            if tc['rules_edit']:
                st.markdown("---")
                st.markdown("#### 检查规则（可编辑）")
                st.caption("修改规则后点击「重新检查」，将按新规则重新评估")
                edited_rules = _render_rules_panel(current_rules, editable=True)
                recheck_count = st.session_state.get('recheck_count', 0)
                recheck_limit = tc['recheck_limit']
                can_recheck = (recheck_limit == -1 or recheck_count < recheck_limit)
                if can_recheck:
                    if st.button("按修改后的规则重新检查", type="primary"):
                        st.session_state['recheck_count'] = recheck_count + 1
                        st.session_state['custom_rules'] = edited_rules
                        st.session_state.pop('check_cache', None)
                        st.rerun()
                else:
                    st.button("按修改后的规则重新检查", disabled=True)
                    st.warning(f"已用完 {recheck_limit} 次复查机会，升级套餐可获得更多复查次数")
            elif tc['rules_view']:
                st.markdown("---")
                st.markdown("#### 本次使用的检查规则")
                _render_rules_panel(current_rules, editable=False)
                recheck_count = st.session_state.get('recheck_count', 0)
                recheck_limit = tc['recheck_limit']
                if recheck_limit > 0:
                    remaining = max(0, recheck_limit - recheck_count)
                    st.caption(f"剩余复查次数：{remaining}/{recheck_limit}")

            # ---- 筛选器（基础版及以上）----
            filtered = issues
            if tc['show_filter']:
                f1, f2, f3 = st.columns(3)
                with f1:
                    sev_f = st.selectbox("严重度", ['全部','错误','警告','建议'], key='sf')
                with f2:
                    mod_f = st.selectbox("模块", ['全部']+[m['name'] for m in mods], key='mf')
                with f3:
                    src_f = st.selectbox("来源", ['全部','官方规定','专业补充','批注修订'], key='rf')
                sev_map = {'错误':'error','警告':'warning','建议':'info'}
                src_map = {'官方规定':'official','专业补充':'supplement','批注修订':'annotation'}
                if sev_f != '全部': filtered = [i for i in filtered if i['severity'] == sev_map[sev_f]]
                if mod_f != '全部': filtered = [i for i in filtered if i['module'] == mod_f]
                if src_f != '全部': filtered = [i for i in filtered if i['source'] == src_map[src_f]]

            # ---- 问题列表（按套餐限制条数）----
            issue_limit = tc['issue_limit']
            display_issues = filtered if issue_limit == -1 else filtered[:issue_limit]
            st.caption(f"显示 {len(display_issues)} / {len(issues)} 条")
            for issue in display_issues:
                st.markdown(render_issue(issue, show_suggestion=tc['show_suggestion']), unsafe_allow_html=True)

            # 极简版：显示剩余条数的升级提示
            if issue_limit != -1 and len(filtered) > issue_limit:
                remaining_count = len(filtered) - issue_limit
                st.markdown(f'''
                <div style="text-align:center;padding:24px;background:var(--bg-card);border-radius:12px;
                    border:1px dashed var(--border-card);margin:12px 0;">
                    <div style="font-size:1.1rem;color:var(--text-primary);margin-bottom:6px;">
                        还有 {remaining_count} 条问题未展示</div>
                    <div style="font-size:0.85rem;color:var(--text-secondary);">
                        升级基础版查看全部问题 + 下载完整报告</div>
                </div>''', unsafe_allow_html=True)

            # 基础版无修改建议时的升级提示
            if not tc['show_suggestion']:
                st.info("升级专业版可查看每条问题的修改建议，快速定位修改方向")

            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

            # ---- 下载按钮（基础版及以上）----
            if tc['can_download']:
                st.download_button("下载完整 HTML 报告", data=html_content,
                    file_name=f"格式审查报告_{report_id}.html", mime="text/html",
                    type="primary", use_container_width=True)
            else:
                st.button("🔒 下载报告（基础版及以上）", disabled=True, use_container_width=True)

            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

            # ---- 一键修复区域 ----
            if tc.get('auto_fix'):
                st.markdown("#### 一键修复")
                from thesis_fixer import ThesisFixer, FIXABLE_MODULES, UNFIXABLE_REASONS

                # 统计可修复项
                fixable_issues = [i for i in data['issues']
                    if i['module'] in FIXABLE_MODULES and ('字体' in i['rule'] or '字号' in i['rule']
                        or '行距' in i['rule'] or '缩进' in i['rule'] or '居中' in i['rule']
                        or '加粗' in i['rule'] or '页边距' in i['rule'] or '纸张' in i['rule'])]
                unfixable_issues = [i for i in data['issues'] if i not in fixable_issues]

                col_f1, col_f2 = st.columns(2)
                col_f1.metric("可自动修复", f"{len(fixable_issues)} 项", help="字体、字号、行距、缩进、对齐等格式问题")
                col_f2.metric("需手动处理", f"{len(unfixable_issues)} 项", help="页眉页脚、页码、编号、内容类问题")

                if len(fixable_issues) == 0:
                    st.info("当前论文没有可自动修复的格式问题，太棒了！")
                else:
                    with st.expander("查看修复预览", expanded=False):
                        st.markdown("**将修复的项目：**")
                        for fi in fixable_issues[:20]:
                            st.markdown(f"- {fi['module']} | {fi['location']} | {fi['rule']}")
                        if len(fixable_issues) > 20:
                            st.caption(f"... 还有 {len(fixable_issues)-20} 项")
                        st.markdown("**不修复的项目（需手动处理）：**")
                        skip_modules = set(i['module'] for i in unfixable_issues if i['module'] in UNFIXABLE_REASONS)
                        for mod in skip_modules:
                            st.markdown(f"- {mod}: {UNFIXABLE_REASONS[mod]}")

                    if st.button("确认修复并下载", type="primary", use_container_width=True):
                        with st.spinner("正在修复格式..."):
                            # 从缓存里的字节重建临时文件（不依赖已被 finally 清理的原 tmp_path）
                            cache = st.session_state.get('check_cache', {})
                            orig_bytes = cache.get('file_bytes')
                            if not orig_bytes:
                                # 场景：用户付款后关闭浏览器→第二天回来→解锁状态还在但 session_state 丢了
                                # 不让用户以为付款打水漂，明确指引 + 保持解锁
                                st.warning(
                                    "⚠ 会话已过期，请重新上传同一份论文即可继续使用修复功能。"
                                    "\n\n你的解锁权限已保留，无需再次付款。")
                            else:
                                import tempfile, gc
                                orig_tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
                                orig_path = orig_tmp.name
                                orig_tmp.write(orig_bytes)
                                orig_tmp.close()

                                fix_checker = None
                                fixer = None
                                try:
                                    # 用缓存中的检查器重建 issues
                                    _fix_rules = st.session_state.get('custom_rules') or \
                                        get_default_rules('本科' if st.session_state.get('edu_level') == '本科' else '硕士')
                                    fix_checker = ThesisChecker(orig_path,
                                        rules=_fix_rules)
                                    fix_checker.run_all_checks()
                                    fixer = ThesisFixer(orig_path, fix_checker.issues, fix_checker.rules)
                                    fix_log, skip_log = fixer.fix_all()
                                    # 保存到临时文件
                                    fixed_path = os.path.join(tempfile.gettempdir(), f"论文_已修复_{report_id}.docx")
                                    fixer.save(fixed_path)

                                    st.success(f"修复完成！已修复 {len(fix_log)} 项，跳过 {len(skip_log)} 项")

                                    # 修复日志
                                    with st.expander("查看修复日志"):
                                        for m, loc, desc in fix_log:
                                            st.markdown(f"✅ {m} | {loc} | {desc}")
                                        if skip_log:
                                            st.markdown("---")
                                            for m, loc, reason in skip_log[:10]:
                                                st.markdown(f"⏭ {m} | {loc} | {reason}")

                                    # 下载修复后的文件
                                    with open(fixed_path, 'rb') as f:
                                        st.download_button(
                                            "📥 下载修复后的论文",
                                            data=f.read(),
                                            file_name=f"论文_已修复.docx",
                                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                            type="primary", use_container_width=True)

                                    st.info("建议下载后重新上传进行复查，确认修复效果")
                                finally:
                                    # Windows 上 python-docx 不是上下文管理器，必须显式释放
                                    # 否则 os.unlink 会抛 PermissionError (WinError 32) 导致临时文件泄漏
                                    try:
                                        if fixer is not None:
                                            fixer.doc = None
                                        if fix_checker is not None:
                                            fix_checker.doc = None
                                        del fixer, fix_checker
                                        gc.collect()
                                    except Exception:
                                        pass
                                    try:
                                        if os.path.exists(orig_path):
                                            os.unlink(orig_path)
                                    except Exception:
                                        pass
            else:
                # 非修复版用户 → 引导升级
                st.markdown('''<div class="glass-card" style="text-align:center;padding:20px;">
                    <div style="font-size:1.1rem;font-weight:700;margin-bottom:8px;">一键修复格式问题</div>
                    <div style="color:var(--text-secondary);font-size:0.85rem;margin-bottom:12px;">
                        升级修复版，自动修复字体、字号、行距、缩进等格式问题<br>
                        人工改格式 300-500 元，修复版仅 99.9 元</div>
                    <div style="color:var(--accent-yellow);font-size:0.8rem;">🔒 修复版专属功能</div>
                </div>''', unsafe_allow_html=True)

    # 页脚
    st.markdown(f'''<div class="app-footer">
        论文格式一键体检 &nbsp;|&nbsp; 报告编号 {report_id} &nbsp;|&nbsp;
        {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        <br>联系微信 l8811925
    </div>''', unsafe_allow_html=True)
    _render_admin_panel()

else:
    # ========== 未上传 - 介绍页 ==========
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # 模块介绍用卡片网格
    modules_names = [
        "页面设置", "封面格式", "摘要规范", "目录格式", "正文格式",
        "标题层级", "图表规范", "页眉页脚", "页码", "参考文献",
        "结构完整", "编号规范", "单位符号", "内容规范",
    ]

    st.markdown("#### 14 个检测模块全覆盖")
    tags = '<div style="display:flex;flex-wrap:wrap;justify-content:center;gap:8px;margin:20px 0;">'
    for name in modules_names:
        tags += f'<span class="module-tag">{name}</span>'
    tags += '</div>'
    st.markdown(tags, unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # 套餐对比
    st.markdown("#### 毕业季特惠 · 套餐说明")
    st.markdown('''
    <div class="pricing-grid">
        <div class="glass-card tier-free" style="text-align:center;padding:24px 16px;">
            <div style="font-size:1.1rem;font-weight:700;margin-bottom:4px;">免费体验版</div>
            <div class="price" style="font-size:2rem;font-weight:800;margin:12px 0;">0 元</div>
            <div style="font-size:0.8rem;color:#94a3b8;line-height:2;">
                总分 + 14个模块评分概览<br>
                免费查看 3 条格式错误详情<br>
                限 2 次检查机会
            </div>
        </div>
        <div class="glass-card tier-free" style="text-align:center;padding:24px 16px;">
            <div style="font-size:1.1rem;font-weight:700;margin-bottom:4px;">极简版</div>
            <div style="margin:12px 0;">
                <span class="original-price">原价 19.9 元</span><br>
                <span class="price" style="font-size:2rem;font-weight:800;">9.9 元</span>
                <span class="discount-badge" style="margin-left:6px;">毕业季半价</span>
            </div>
            <div style="font-size:0.8rem;color:#94a3b8;line-height:2;text-align:left;padding:0 12px;">
                <i style="color:#cbd5e1;">要一份能看的清单，自己改</i><br>
                14 模块 + 30+ 格式雷区扫描<br>
                查看前 5 条问题详情<br>
                含问题位置定位 · 1 次检查
            </div>
        </div>
        <div class="glass-card tier-basic" style="text-align:center;padding:24px 16px;">
            <div style="font-size:1.1rem;font-weight:700;margin-bottom:4px;">基础版</div>
            <div style="margin:12px 0;">
                <span class="original-price">原价 49.9 元</span><br>
                <span class="price" style="font-size:2rem;font-weight:800;">24.9 元</span>
                <span class="discount-badge" style="margin-left:6px;">毕业季5折</span>
            </div>
            <div style="font-size:0.8rem;color:#94a3b8;line-height:2;text-align:left;padding:0 12px;">
                <i style="color:#cbd5e1;">清单 + 每条问题修改示范</i><br>
                极简版全部功能<br>
                <b>查看全部问题详情</b> · 筛选<br>
                下载完整 HTML 报告<br>
                3 次复查（初稿+修改稿+终稿）
            </div>
        </div>
        <div class="glass-card tier-pro" style="text-align:center;padding:24px 16px;">
            <div style="font-size:1.1rem;font-weight:700;margin-bottom:4px;">专业版 <span class="recommend-badge">推荐</span></div>
            <div style="margin:12px 0;">
                <span class="original-price">原价 99.9 元</span><br>
                <span class="price" style="font-size:2rem;font-weight:800;">49.9 元</span>
                <span class="discount-badge" style="margin-left:6px;">毕业季5折</span>
            </div>
            <div style="font-size:0.8rem;color:#94a3b8;line-height:2;text-align:left;padding:0 12px;">
                <i style="color:#cbd5e1;">定位到页码行号，适合反复改 3 版以上</i><br>
                基础版全部功能<br>
                <b>每条问题附修改建议</b><br>
                <b>自定义编辑全部检查规则</b><br>
                适配任意学校 · 不限次复查 · 优先客服
            </div>
        </div>
    </div>
    <div style="max-width:360px;margin:16px auto;">
        <div class="glass-card tier-pro" style="text-align:center;padding:24px 16px;border-color:rgba(234,179,8,0.6);">
            <div style="font-size:1.1rem;font-weight:700;margin-bottom:4px;">修复版 <span class="recommend-badge" style="background:linear-gradient(135deg,#f59e0b,#ef4444);">省300元</span></div>
            <div style="margin:12px 0;">
                <span class="original-price">原价 199.9 元</span><br>
                <span class="price" style="font-size:2rem;font-weight:800;">99.9 元</span>
                <span class="discount-badge" style="margin-left:6px;">毕业季5折</span>
            </div>
            <div style="font-size:0.8rem;color:#94a3b8;line-height:2;text-align:left;padding:0 12px;">
                <i style="color:#cbd5e1;">直接给你改好的 docx · 省 300-500 元人工</i><br>
                专业版全部功能<br>
                <b>一键自动修复格式问题</b><br>
                <b>下载修复后的论文文件</b><br>
                修复前预览 · 修复后复查
            </div>
            <div style="font-size:0.72rem;color:#f59e0b;margin-top:8px;">人工代改 300-500 元 · 本工具 99.9 元 · 周末解放</div>
        </div>
    </div>
    <div style="text-align:center;font-size:0.8rem;color:#64748b;margin-bottom:20px;">
        邀请同学使用你的专属邀请码购买，双方各返 <b style="color:#818cf8;">5 元</b>（付费后自动获得邀请码）
    </div>
    ''', unsafe_allow_html=True)

    # 朋友圈 / 师门群可直接转发的分享语
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('#### 觉得有用？转给还在熬夜调格式的同学')
    st.markdown('''
    <div style="display:flex;flex-direction:column;gap:10px;margin:12px 0 20px;">
        <div style="padding:12px 16px;background:rgba(99,102,241,0.08);border-left:3px solid #6366f1;border-radius:8px;color:var(--text-secondary);font-size:0.88rem;">
            "查完才知道，我那篇'改了八遍'的论文，还有 60 多处格式不合规。"
        </div>
        <div style="padding:12px 16px;background:rgba(34,197,94,0.08);border-left:3px solid #22c55e;border-radius:8px;color:var(--text-secondary);font-size:0.88rem;">
            "人工改格式报价 400，这个 99 块直接给我一份改好的 docx，真香。"
        </div>
        <div style="padding:12px 16px;background:rgba(234,179,8,0.08);border-left:3px solid #eab308;border-radius:8px;color:var(--text-secondary);font-size:0.88rem;">
            "不敢说它全能，但答辩老师常盯的那几项，它是真的都查了。"
        </div>
        <div style="padding:12px 16px;background:rgba(239,68,68,0.08);border-left:3px solid #ef4444;border-radius:8px;color:var(--text-secondary);font-size:0.88rem;">
            "导师只会说'格式自己弄一下'，它会告诉你第 23 页第 4 行空了两格。"
        </div>
        <div style="padding:12px 16px;background:rgba(59,130,246,0.08);border-left:3px solid #3b82f6;border-radius:8px;color:var(--text-secondary);font-size:0.88rem;">
            "转给还在熬夜调页眉的学弟学妹 / 师弟师妹，别再手动数空行了。"
        </div>
    </div>
    <div style="text-align:center;font-size:0.78rem;color:var(--text-muted);margin-bottom:16px;">
        直接复制上面任意一句，粘到朋友圈 / 毕业群 / 师门群
    </div>
    ''', unsafe_allow_html=True)

    # 底部
    st.markdown('''<div class="app-footer">
        论文格式一键体检 &nbsp;|&nbsp; 本科 / 研究生毕业论文通用 &nbsp;|&nbsp; 联系微信 l8811925
        <br>人工改格式 300-500 元，用工具最低 9.9 元
        <br>Beta 版 · 覆盖毕业论文常见 30+ 格式雷区 · 报告不对无理由退款
    </div>''', unsafe_allow_html=True)
    _render_admin_panel()
