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
import string
import random
from datetime import datetime
from thesis_checker import ThesisChecker

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
    content:"仅支持 .docx 格式，最大 50MB"; font-size:0.75rem; color:var(--text-muted);
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
.pricing-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:16px; margin-bottom:12px; }
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
# 兑换码管理（文件锁 + 会话绑定）
# ============================================================
CODES_FILE = os.path.join(os.path.dirname(__file__), 'codes.json')
LOCK_FILE = CODES_FILE + '.lock'
ADMIN_PWD = "8811925123Aa!"

import uuid
if os.name != 'nt':
    import fcntl
else:
    import msvcrt

def _get_session_id():
    """每个浏览器会话生成唯一 ID（存在 session_state 中，刷新不变）"""
    if 'session_id' not in st.session_state:
        st.session_state['session_id'] = str(uuid.uuid4())[:8]
    return st.session_state['session_id']

def _locked_read_write(fn):
    """带文件锁的读写操作，防止多用户并发写入冲突"""
    import functools
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        if os.name == 'nt':
            # Windows: 用临时锁文件 + 重试
            for _ in range(10):
                try:
                    lock_fd = open(LOCK_FILE, 'w')
                    msvcrt.locking(lock_fd.fileno(), 1, 1)  # LK_NBLCK
                    try:
                        return fn(*args, **kwargs)
                    finally:
                        msvcrt.locking(lock_fd.fileno(), 0, 1)  # unlock
                        lock_fd.close()
                except (OSError, IOError):
                    time.sleep(0.1)
            return fn(*args, **kwargs)  # 超时兜底
        else:
            # Linux/Mac: fcntl 文件锁
            lock_fd = open(LOCK_FILE, 'w')
            try:
                fcntl.flock(lock_fd, fcntl.LOCK_EX)
                return fn(*args, **kwargs)
            finally:
                fcntl.flock(lock_fd, fcntl.LOCK_UN)
                lock_fd.close()
    return wrapper

def load_codes():
    if os.path.exists(CODES_FILE):
        with open(CODES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_codes(codes):
    with open(CODES_FILE, 'w', encoding='utf-8') as f:
        json.dump(codes, f, ensure_ascii=False, indent=2)

@_locked_read_write
def verify_code(code, report_id=None, filename=None):
    """验证兑换码并绑定到当前会话和报告"""
    codes = load_codes()
    code = code.strip().upper()
    if code in codes:
        if codes[code]['used']:
            return False, '此兑换码已被使用'
        codes[code]['used'] = True
        codes[code]['used_at'] = datetime.now().isoformat()
        codes[code]['session'] = _get_session_id()
        if report_id:
            codes[code]['report_id'] = report_id
        if filename:
            codes[code]['filename'] = filename
        save_codes(codes)
        return True, '解锁成功'
    return False, '兑换码无效'

@_locked_read_write
def generate_codes(n=20, tier='basic'):
    codes = load_codes()
    new_codes = []
    for _ in range(n):
        code = f"FMT-{''.join(random.choices(string.ascii_uppercase + string.digits, k=4))}-{''.join(random.choices(string.ascii_uppercase + string.digits, k=4))}"
        if code not in codes:
            codes[code] = {'tier': tier, 'used': False, 'created': datetime.now().isoformat()}
            new_codes.append(code)
    save_codes(codes)
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
def render_issue(issue):
    sev_c = {'error':'#f87171','warning':'#fbbf24','info':'#60a5fa'}
    src_c = {'official':'#a78bfa','supplement':'#2dd4bf','annotation':'#fb923c'}
    bc = sev_c.get(issue['severity'],'#64748b')
    sc = src_c.get(issue['source'],'#2dd4bf')
    preview = ''
    if issue.get('text_preview') and issue['text_preview'] != '(空)':
        preview = f'<div style="font-size:0.75rem;color:#64748b;margin-top:4px;">{issue["text_preview"]}</div>'
    return f'''<div class="issue-card" style="border-left:3px solid {bc};">
      <div style="display:flex;gap:8px;margin-bottom:6px;flex-wrap:wrap;">
        <span style="background:{bc}22;color:{bc};padding:2px 10px;border-radius:4px;font-size:0.75rem;font-weight:600;">{issue['severity_label']}</span>
        <span style="background:#33415522;color:#94a3b8;padding:2px 10px;border-radius:4px;font-size:0.75rem;">{issue['module']}</span>
        <span style="background:{sc}22;color:{sc};padding:2px 10px;border-radius:4px;font-size:0.75rem;font-weight:600;">{issue['source_label']}</span>
        <span style="color:#64748b;font-size:0.75rem;font-family:monospace;">{issue['location']}</span>
      </div>
      <div style="font-size:0.9rem;color:#e2e8f0;margin-bottom:4px;">{issue['rule']}</div>
      <div style="font-size:0.8rem;">
        <span style="color:#10b981;">期望: {issue['expected']}</span> &nbsp;→&nbsp;
        <span style="color:#ef4444;">实际: {issue['actual']}</span>
      </div>{preview}
    </div>'''

# ============================================================
# 渲染模块卡片
# ============================================================
def render_module_card(mod):
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
                gen_tier = st.selectbox("套餐", ['basic', 'pro'])
            if st.button("生成兑换码", use_container_width=True):
                new_codes = generate_codes(gen_n, gen_tier)
                st.code('\n'.join(new_codes))
            codes = load_codes()
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

# Hero
st.markdown('''
<div class="hero">
    <h1>论文格式一键体检</h1>
    <p>别让格式问题成为你毕业的最后一道坎 —— 上传论文，快速找出所有排版问题</p>
    <div class="highlights">
        <div class="hl-item"><div class="hl-num" style="color:#3b82f6;">60+</div><div class="hl-label">检查规则</div></div>
        <div class="hl-item"><div class="hl-num" style="color:#3b82f6;">13</div><div class="hl-label">检测模块</div></div>
        <div class="hl-item"><div class="hl-num" style="color:#3b82f6;">省90%</div><div class="hl-label">对比人工改格式</div></div>
        <div class="hl-item"><div class="hl-num" style="color:#3b82f6;">2,400+</div><div class="hl-label">已完成检测</div></div>
    </div>
</div>
''', unsafe_allow_html=True)

# 上传区
st.markdown('<div class="upload-zone">', unsafe_allow_html=True)
col_up, col_title = st.columns([3, 2])
with col_up:
    uploaded_file = st.file_uploader("上传论文 (.docx)", type=['docx'],
        help="支持 .docx 格式，最大 200MB", label_visibility="collapsed")
with col_title:
    thesis_title = st.text_input("论文题目（可选，用于页眉校验）",
        placeholder="如：基于深度学习的小麦病害图像识别研究",
        label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center;font-size:0.9rem;color:#94a3b8;margin:8px 0 16px;">已为 2,400+ 篇论文完成格式体检 · 检测不准确全额退款</p>', unsafe_allow_html=True)

# ============================================================
# 免费次数限制（浏览器指纹 + session_state）
# ============================================================
FREE_CHECK_LIMIT = 2  # 免费版最多检查 2 次

def _get_usage_count():
    """获取当前会话的免费使用次数"""
    return st.session_state.get('free_usage', 0)

def _increment_usage():
    st.session_state['free_usage'] = st.session_state.get('free_usage', 0) + 1

# 注入 JS：用 localStorage 跨会话持久化计数
# 页面加载时读取 localStorage，写入隐藏的 query param
import streamlit.components.v1 as components

_COUNTER_JS = """
<script>
(function(){
    var key = 'fmt_free_count';
    var count = parseInt(localStorage.getItem(key) || '0');
    // 把计数写入 iframe 的 title，供 Streamlit 读取
    document.title = count.toString();

    // 监听来自 Streamlit 的递增指令
    window.addEventListener('message', function(e){
        if(e.data && e.data.type === 'fmt_increment'){
            count = parseInt(localStorage.getItem(key) || '0') + 1;
            localStorage.setItem(key, count.toString());
        }
    });
})();
</script>
"""

# 在检查时递增 localStorage 计数
_INCREMENT_JS = """
<script>
(function(){
    var key = 'fmt_free_count';
    var count = parseInt(localStorage.getItem(key) || '0') + 1;
    localStorage.setItem(key, count.toString());
})();
</script>
"""

# ============================================================
# 检查流程
# ============================================================
_can_check = False
if uploaded_file is not None:
    usage = _get_usage_count()
    unlocked_session = st.session_state.get('unlocked', False)
    in_payment_flow = 'pay_tier' in st.session_state or 'auto_code' in st.session_state
    if usage >= FREE_CHECK_LIMIT and not unlocked_session and not in_payment_flow:
        st.warning(f"免费版已用完 {FREE_CHECK_LIMIT} 次检查机会，请购买套餐后使用兑换码解锁")
    else:
        _can_check = True

if uploaded_file is not None and _can_check:
    # 用文件内容的 hash 判断是否同一份论文，避免 rerun 时重复检查
    _file_bytes = uploaded_file.getvalue()
    _file_hash = hashlib.md5(_file_bytes).hexdigest()
    _cache = st.session_state.get('check_cache', {})
    _cache_hit = (_cache.get('file_hash') == _file_hash
                  and _cache.get('thesis_title') == (thesis_title or None))

    if _cache_hit:
        # 缓存命中：直接读取上次结果，不重新检查
        data = _cache['data']
        html_content = _cache['html_content']
        report_id = _cache['report_id']
    else:
        # 首次检查或文件变化：执行检查并缓存
        # 递增免费计数（仅免费用户，付款流程中不再递增）
        if not st.session_state.get('unlocked', False) and 'pay_tier' not in st.session_state and 'auto_code' not in st.session_state:
            _increment_usage()
            components.html(_INCREMENT_JS, height=0)

        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            tmp.write(_file_bytes)
            tmp_path = tmp.name

        html_path = None
        try:
            with st.spinner("正在审查论文格式..."):
                t0 = time.time()
                checker = ThesisChecker(tmp_path, thesis_title=thesis_title or None)
                checker.run_all_checks()
                elapsed = time.time() - t0
                data = checker.get_report_data()
                html_path = tmp_path.replace('.docx', '_report.html')
                checker.generate_html_report(html_path)
                with open(html_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
        finally:
            try:
                os.unlink(tmp_path)
                if html_path and os.path.exists(html_path):
                    os.unlink(html_path)
            except:
                pass

        report_id = f"FMT-{datetime.now().strftime('%Y%m%d')}-{hashlib.md5(_file_bytes[:1024]).hexdigest()[:6].upper()}"

        # 写入缓存
        st.session_state['check_cache'] = {
            'file_hash': _file_hash,
            'thesis_title': thesis_title or None,
            'data': data,
            'html_content': html_content,
            'report_id': report_id,
        }

    st.success(f"审查完成！报告编号 {report_id}")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ========== 评分概览 ==========
    col_ring, col_metrics = st.columns([1, 2])
    with col_ring:
        st.markdown(render_score_ring(data['total_score'], data['max_score'], data['grade']),
            unsafe_allow_html=True)
        # 排名机制（基于分数估算百分位）
        pct_score = data['pct']
        if pct_score >= 90: beat_pct = 95
        elif pct_score >= 80: beat_pct = 80
        elif pct_score >= 70: beat_pct = 55
        elif pct_score >= 60: beat_pct = 35
        elif pct_score >= 50: beat_pct = 20
        else: beat_pct = 8
        st.markdown(f'<p style="text-align:center;font-size:0.85rem;color:#94a3b8;margin-top:4px;">'
            f'你的论文格式超过了 <b style="color:#818cf8;">{beat_pct}%</b> 的论文</p>',
            unsafe_allow_html=True)
    with col_metrics:
        m1, m2, m3 = st.columns(3)
        m1.metric("严重错误", data['error_count'])
        m2.metric("格式警告", data['warning_count'])
        m3.metric("优化建议", data['info_count'])
        m4, m5, m6 = st.columns(3)
        m4.metric("段落数", data['total_paras'])
        m5.metric("表格数", data['total_tables'])
        m6.metric("图片数", data['total_images'])

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ========== 模块评分卡片网格 ==========
    st.markdown("#### 各模块得分")
    mods = data['modules']
    # 用 HTML grid 渲染（比 st.columns 更紧凑）
    grid_html = '<div class="mod-grid">'
    for mod in mods:
        grid_html += render_module_card(mod)
    grid_html += '</div>'
    st.markdown(grid_html, unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # ========== 问题列表 ==========
    issues = data['issues']
    st.markdown(f"#### 问题详情（共 {len(issues)} 条）")

    # 免费展示：优先挑选编号规范和正文格式的 error/warning（最抓眼球）
    FREE_LIMIT = 3
    priority_modules = ['编号规范', '正文格式', '标题层级', '图表规范']
    priority_issues = [i for i in issues
                       if i['module'] in priority_modules and i['severity'] in ('error', 'warning')]
    other_issues = [i for i in issues if i not in priority_issues]
    free_preview = (priority_issues + other_issues)[:FREE_LIMIT]

    for issue in free_preview:
        st.markdown(render_issue(issue), unsafe_allow_html=True)

    # ========== 付费墙 / 完整报告 ==========
    if len(issues) > FREE_LIMIT:
        unlocked = st.session_state.get('unlocked', False)

        if not unlocked:
            # 付费墙遮罩
            st.markdown(f'''
            <div class="paywall">
                <div style="font-size:2.5rem;margin-bottom:8px;">🔒</div>
                <div style="font-size:1.2rem;font-weight:700;color:var(--text-primary);margin-bottom:6px;">
                    还有 {len(issues)-FREE_LIMIT} 条问题待查看</div>
                <div style="color:var(--text-secondary);font-size:0.9rem;margin-bottom:20px;">
                    选择套餐解锁完整报告，查看全部问题的位置和修改建议</div>
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
                        60+ 项规则全量扫描<br>
                        全部问题列表 + 位置定位<br>
                        单次使用</div>
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
                        极简版全部功能<br>
                        每条问题附修改建议<br>
                        按严重度/模块智能筛选<br>
                        下载完整 HTML 报告<br>
                        含 3 次复查</div>
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
                        基础版全部功能<br>
                        支持自定义学校专属审查模板<br>
                        不限次复查 60 天<br>
                        赠送 Word 排版模板<br>
                        优先客服响应</div>
                </div>''', unsafe_allow_html=True)
                pick_pro = st.button("选择专业版", key="pick_pro", type="primary", use_container_width=True)

            st.caption("邀请同学使用你的专属邀请码购买，双方各返 5 元")

            # 选定套餐后弹出付款区（用按钮当前帧判断，不持久化到 session_state）
            just_picked = None
            if pick_lite: just_picked = ("极简版", "9.9")
            elif pick_basic: just_picked = ("基础版", "24.9")
            elif pick_pro: just_picked = ("专业版", "49.9")

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
                    pay_img = os.path.join(os.path.dirname(__file__), 'zhifubao.jpg')
                    if os.path.exists(pay_img):
                        st.image(pay_img, width=200, caption=f"支付宝扫码 · {tier_price}元")
                    else:
                        st.info(f"请扫码支付 {tier_price} 元")

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
                        解锁后包含：<br>
                        &nbsp;&nbsp;全部问题的详细位置和修改建议<br>
                        &nbsp;&nbsp;按严重度 / 模块 / 来源筛选<br>
                        &nbsp;&nbsp;下载完整 HTML 报告文件
                    </div>""", unsafe_allow_html=True)

        else:
            # ========== 完整报告 ==========
            st.markdown("**已解锁完整报告**")

            f1, f2, f3 = st.columns(3)
            with f1:
                sev_f = st.selectbox("严重度", ['全部','错误','警告','建议'], key='sf')
            with f2:
                mod_f = st.selectbox("模块", ['全部']+[m['name'] for m in mods], key='mf')
            with f3:
                src_f = st.selectbox("来源", ['全部','官方规定','专业补充','批注修订'], key='rf')

            sev_map = {'错误':'error','警告':'warning','建议':'info'}
            src_map = {'官方规定':'official','专业补充':'supplement','批注修订':'annotation'}
            filtered = issues
            if sev_f != '全部': filtered = [i for i in filtered if i['severity'] == sev_map[sev_f]]
            if mod_f != '全部': filtered = [i for i in filtered if i['module'] == mod_f]
            if src_f != '全部': filtered = [i for i in filtered if i['source'] == src_map[src_f]]

            st.caption(f"显示 {len(filtered)} / {len(issues)} 条")
            for issue in filtered:
                st.markdown(render_issue(issue), unsafe_allow_html=True)

            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            st.download_button("下载完整 HTML 报告", data=html_content,
                file_name=f"格式审查报告_{report_id}.html", mime="text/html",
                type="primary", use_container_width=True)

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
        "标题层级", "图表规范", "页眉页脚", "参考文献", "结构完整",
        "编号规范", "单位符号", "内容规范",
    ]

    st.markdown("#### 13 个检测模块全覆盖")
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
                总分 + 13个模块评分概览<br>
                免费查看 3 条格式错误详情<br>
                限 2 次检查机会
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
                60+ 项格式规则全量扫描<br>
                全部问题精确到段落 + 修改建议<br>
                按严重度/模块/来源智能筛选<br>
                下载完整 HTML 审查报告<br>
                含 3 次复查（初稿+修改稿+终稿）
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
                基础版全部功能<br>
                支持自定义学校专属审查模板<br>
                不限次复查 60 天（改到毕业）<br>
                赠送 Word 排版模板 (.dotx)<br>
                优先客服响应
            </div>
        </div>
    </div>
    <div style="text-align:center;font-size:0.8rem;color:#64748b;margin-bottom:20px;">
        邀请同学使用你的专属邀请码购买，双方各返 <b style="color:#818cf8;">5 元</b>（付费后自动获得邀请码）
    </div>
    ''', unsafe_allow_html=True)

    # 底部
    st.markdown('''<div class="app-footer">
        论文格式一键体检 &nbsp;|&nbsp; 联系微信 l8811925
        <br>人工改格式 300-500 元，用工具最低 9.9 元，省 95%+
        <br>检测不准确全额退款
    </div>''', unsafe_allow_html=True)
    _render_admin_panel()
