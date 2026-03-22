"""
论文格式智能审查平台 - Streamlit Web 应用
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
    page_title="论文格式智能审查",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 自定义 CSS
# ============================================================
st.markdown("""
<style>
/* 顶部渐变条 */
.stApp > header { display: none; }
div[data-testid="stAppViewContainer"]::before {
    content: '';
    display: block;
    height: 3px;
    background: linear-gradient(90deg, #3b82f6, #8b5cf6, #ec4899, #f59e0b);
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    z-index: 9999;
}

/* 卡片样式 */
div[data-testid="stMetric"] {
    background: rgba(26,35,50,0.7);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 12px;
    padding: 16px;
    backdrop-filter: blur(12px);
}

/* 表格美化 */
.stDataFrame { border-radius: 8px; overflow: hidden; }

/* 隐藏 streamlit 底部 */
footer { visibility: hidden; }

/* 付费墙 */
.paywall-overlay {
    background: linear-gradient(180deg, transparent, rgba(12,18,34,0.85) 40%, rgba(12,18,34,0.98));
    padding: 60px 20px 40px;
    text-align: center;
    border-radius: 12px;
    margin-top: -80px;
    position: relative;
    z-index: 10;
}
.paywall-btn {
    display: inline-block;
    background: linear-gradient(135deg, #f59e0b, #f97316);
    color: #000;
    font-weight: 700;
    padding: 14px 48px;
    border-radius: 10px;
    font-size: 1.1rem;
    text-decoration: none;
    box-shadow: 0 4px 20px rgba(245,158,11,0.35);
}

/* SVG 环形图居中 */
.score-ring { text-align: center; }

/* 问题行颜色 */
.severity-error { color: #f87171; }
.severity-warning { color: #fbbf24; }
.severity-info { color: #60a5fa; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 兑换码管理
# ============================================================
CODES_FILE = os.path.join(os.path.dirname(__file__), 'codes.json')

def load_codes():
    if os.path.exists(CODES_FILE):
        with open(CODES_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_codes(codes):
    with open(CODES_FILE, 'w') as f:
        json.dump(codes, f, ensure_ascii=False, indent=2)

def verify_code(code):
    codes = load_codes()
    code = code.strip().upper()
    if code in codes:
        if codes[code]['used']:
            return False, '此兑换码已被使用'
        codes[code]['used'] = True
        codes[code]['used_at'] = datetime.now().isoformat()
        save_codes(codes)
        return True, '解锁成功'
    return False, '兑换码无效'

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
# SVG 环形评分图
# ============================================================
def render_score_ring(score, max_score, grade):
    pct = score / max_score * 100 if max_score > 0 else 0
    radius = 70
    circumference = 2 * 3.14159 * radius
    offset = circumference - (circumference * pct / 100)

    if pct >= 80: colors = ('#10b981', '#34d399', 'rgba(16,185,129,0.15)')
    elif pct >= 60: colors = ('#3b82f6', '#60a5fa', 'rgba(59,130,246,0.15)')
    elif pct >= 40: colors = ('#f59e0b', '#fbbf24', 'rgba(245,158,11,0.15)')
    else: colors = ('#ef4444', '#f87171', 'rgba(239,68,68,0.15)')

    grade_labels = {'A': '优秀', 'B': '良好', 'C': '中等', 'D': '及格', 'F': '不及格'}

    svg = f'''
    <div class="score-ring">
    <svg width="200" height="200" viewBox="0 0 200 200">
        <defs>
            <linearGradient id="scoreGrad" x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" style="stop-color:{colors[0]}"/>
                <stop offset="100%" style="stop-color:{colors[1]}"/>
            </linearGradient>
        </defs>
        <circle cx="100" cy="100" r="{radius}" fill="none" stroke="rgba(255,255,255,0.08)" stroke-width="10"/>
        <circle cx="100" cy="100" r="{radius}" fill="none" stroke="url(#scoreGrad)" stroke-width="10"
            stroke-linecap="round" stroke-dasharray="{circumference}" stroke-dashoffset="{offset}"
            transform="rotate(-90 100 100)"
            style="filter: drop-shadow(0 0 8px {colors[0]}40); transition: stroke-dashoffset 1.2s ease-out;"/>
        <text x="100" y="90" text-anchor="middle" fill="#f1f5f9" font-size="40" font-weight="800"
            font-family="system-ui">{score:.0f}</text>
        <text x="100" y="112" text-anchor="middle" fill="#64748b" font-size="14"
            font-family="system-ui">/ {max_score}</text>
        <text x="100" y="140" text-anchor="middle" font-size="14" font-weight="700"
            fill="{colors[0]}" font-family="system-ui">{grade} {grade_labels.get(grade, '')}</text>
    </svg>
    </div>
    '''
    return svg

# ============================================================
# 侧边栏
# ============================================================
with st.sidebar:
    st.markdown("### 📋 论文格式智能审查")
    st.markdown("---")
    st.markdown("""
    **60+ 项规则 | 13 个模块 | 秒出报告**

    覆盖字体字号、页边距、图表编号、
    参考文献、页眉页脚、单位符号等

    ---
    **套餐说明**

    | 版本 | 内容 |
    |------|------|
    | 免费版 | 总分 + 模块评分 + 前3条问题 |
    | 基础版 (29.9元) | 完整报告 + 3次复查 |
    | 专业版 (69.9元) | 基础版 + 不限次复查60天 |

    ---
    """)
    st.markdown("**联系方式**: 微信 `l8811925`")
    st.caption("已帮助 1,200+ 位同学通过格式审查")

    # 管理员入口
    with st.expander("管理员", expanded=False):
        admin_pwd = st.text_input("管理密码", type="password", key="admin_pwd")
        if admin_pwd == "admin2026":
            st.success("管理员已登录")
            if st.button("生成20个兑换码"):
                new_codes = generate_codes(20, 'basic')
                st.code('\n'.join(new_codes))
            codes = load_codes()
            unused = sum(1 for c in codes.values() if not c['used'])
            used = sum(1 for c in codes.values() if c['used'])
            st.metric("未使用", unused)
            st.metric("已使用", used)
            if st.button("查看所有兑换码"):
                for code, info in codes.items():
                    status = "已用" if info['used'] else "可用"
                    st.text(f"{code}  [{status}]  {info.get('tier','basic')}")

# ============================================================
# 主页面
# ============================================================
st.markdown("# 📋 论文格式智能审查")
st.markdown("上传论文 Word 文档，**3秒**获取专业格式审查报告")
st.markdown("---")

# 上传区域
col_upload, col_options = st.columns([2, 1])

with col_upload:
    uploaded_file = st.file_uploader(
        "上传论文 (.docx)",
        type=['docx'],
        help="支持 .docx 格式，文件大小不超过 50MB"
    )

with col_options:
    thesis_title = st.text_input(
        "论文题目（可选，用于页眉校验）",
        placeholder="如：基于深度学习的小麦病害图像识别研究"
    )

# ============================================================
# 检查流程
# ============================================================
if uploaded_file is not None:
    # 保存上传文件到临时目录
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    try:
        # 运行检查
        with st.spinner("正在审查论文格式（13个模块，60+项规则）..."):
            start_time = time.time()
            checker = ThesisChecker(tmp_path, thesis_title=thesis_title or None)
            checker.run_all_checks()
            elapsed = time.time() - start_time
            data = checker.get_report_data()

            # 生成 HTML 报告（供下载）
            html_path = tmp_path.replace('.docx', '_report.html')
            checker.generate_html_report(html_path)
            with open(html_path, 'r', encoding='utf-8') as f:
                html_content = f.read()

        # 生成报告编号
        report_id = f"FMT-{datetime.now().strftime('%Y%m%d')}-{hashlib.md5(uploaded_file.getvalue()[:1024]).hexdigest()[:6].upper()}"

        st.success(f"审查完成! 用时 {elapsed:.1f} 秒 | 报告编号: {report_id}")
        st.markdown("---")

        # ============================
        # 免费部分：概览（所有人可见）
        # ============================

        # 评分区域
        col_score, col_stats = st.columns([1, 2])

        with col_score:
            st.markdown(
                render_score_ring(data['total_score'], data['max_score'], data['grade']),
                unsafe_allow_html=True
            )

        with col_stats:
            c1, c2, c3 = st.columns(3)
            c1.metric("错误", data['error_count'], delta=None)
            c2.metric("警告", data['warning_count'], delta=None)
            c3.metric("建议", data['info_count'], delta=None)

            c4, c5, c6 = st.columns(3)
            c4.metric("段落数", data['total_paras'])
            c5.metric("表格数", data['total_tables'])
            c6.metric("图片数", data['total_images'])

        st.markdown("---")

        # 模块评分卡片
        st.markdown("### 各模块得分")
        cols = st.columns(4)
        for idx, mod in enumerate(data['modules']):
            with cols[idx % 4]:
                pct = mod['pct']
                if pct >= 80: color = '#10b981'
                elif pct >= 60: color = '#3b82f6'
                elif pct >= 40: color = '#f59e0b'
                else: color = '#ef4444'

                st.markdown(f"""
                <div style="background:rgba(26,35,50,0.7); border:1px solid rgba(255,255,255,0.06);
                    border-radius:12px; padding:16px; margin-bottom:12px; backdrop-filter:blur(12px);">
                    <div style="display:flex; justify-content:space-between; margin-bottom:8px;">
                        <span style="font-weight:600; font-size:0.9rem;">{mod['name']}</span>
                        <span style="font-weight:700; color:{color}; font-family:monospace;">{mod['earned']:.1f}/{mod['weight']}</span>
                    </div>
                    <div style="width:100%; height:6px; background:rgba(255,255,255,0.06); border-radius:3px; overflow:hidden;">
                        <div style="width:{pct:.0f}%; height:100%; background:linear-gradient(90deg,{color},{color}aa);
                            border-radius:3px; transition:width 0.8s;"></div>
                    </div>
                    <div style="font-size:0.75rem; color:#64748b; margin-top:6px;">
                        错误:{mod['errors']} 警告:{mod['warnings']}
                    </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("---")

        # ============================
        # 问题列表（免费版显示前3条）
        # ============================
        st.markdown(f"### 问题详情（共 {len(data['issues'])} 条）")

        # 显示前3条（免费）
        FREE_LIMIT = 3
        issues = data['issues']

        for i, issue in enumerate(issues[:FREE_LIMIT]):
            sev_colors = {'error': '#f87171', 'warning': '#fbbf24', 'info': '#60a5fa'}
            src_colors = {'official': '#a78bfa', 'supplement': '#2dd4bf', 'annotation': '#fb923c'}
            border_color = sev_colors.get(issue['severity'], '#64748b')

            st.markdown(f"""
            <div style="border-left:3px solid {border_color}; padding:12px 16px; margin-bottom:8px;
                background:rgba(26,35,50,0.5); border-radius:0 8px 8px 0;">
                <div style="display:flex; gap:8px; margin-bottom:6px; flex-wrap:wrap;">
                    <span style="background:{border_color}22; color:{border_color}; padding:2px 10px;
                        border-radius:4px; font-size:0.75rem; font-weight:600;">{issue['severity_label']}</span>
                    <span style="background:#33415522; color:#94a3b8; padding:2px 10px;
                        border-radius:4px; font-size:0.75rem;">{issue['module']}</span>
                    <span style="background:{src_colors.get(issue['source'], '#2dd4bf')}22;
                        color:{src_colors.get(issue['source'], '#2dd4bf')}; padding:2px 10px;
                        border-radius:4px; font-size:0.75rem; font-weight:600;">{issue['source_label']}</span>
                    <span style="color:#64748b; font-size:0.75rem; font-family:monospace;">{issue['location']}</span>
                </div>
                <div style="font-size:0.9rem; color:#e2e8f0; margin-bottom:4px;">{issue['rule']}</div>
                <div style="font-size:0.8rem;">
                    <span style="color:#10b981;">期望: {issue['expected']}</span> &nbsp;→&nbsp;
                    <span style="color:#ef4444;">实际: {issue['actual']}</span>
                </div>
                {'<div style=\"font-size:0.75rem; color:#64748b; margin-top:4px;\">'+issue["text_preview"]+'</div>' if issue['text_preview'] and issue['text_preview'] != '(空)' else ''}
            </div>
            """, unsafe_allow_html=True)

        # ============================
        # 付费墙
        # ============================
        if len(issues) > FREE_LIMIT:
            # 检查是否已解锁
            unlocked = st.session_state.get('unlocked', False)

            if not unlocked:
                st.markdown(f"""
                <div class="paywall-overlay">
                    <div style="font-size:2rem; margin-bottom:12px;">🔒</div>
                    <div style="font-size:1.3rem; font-weight:700; color:#f1f5f9; margin-bottom:8px;">
                        还有 {len(issues) - FREE_LIMIT} 条问题待查看
                    </div>
                    <div style="font-size:0.95rem; color:#94a3b8; margin-bottom:20px; max-width:500px; margin-left:auto; margin-right:auto;">
                        解锁完整报告，查看所有问题的详细位置和修改建议
                    </div>
                    <div style="display:flex; justify-content:center; gap:24px; margin-bottom:16px; flex-wrap:wrap;">
                        <div style="text-align:center;">
                            <div style="font-size:1.5rem; font-weight:700; color:#f1f5f9;">13</div>
                            <div style="font-size:0.75rem; color:#64748b;">检测模块</div>
                        </div>
                        <div style="text-align:center;">
                            <div style="font-size:1.5rem; font-weight:700; color:#f1f5f9;">60+</div>
                            <div style="font-size:0.75rem; color:#64748b;">检查规则</div>
                        </div>
                        <div style="text-align:center;">
                            <div style="font-size:1.5rem; font-weight:700; color:#f1f5f9;">{elapsed:.1f}s</div>
                            <div style="font-size:0.75rem; color:#64748b;">检测用时</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # 兑换码输入
                st.markdown("---")
                st.markdown("#### 🔑 输入兑换码解锁完整报告")
                st.markdown("付款后联系微信 `l8811925` 获取兑换码")

                col_code, col_btn = st.columns([3, 1])
                with col_code:
                    code_input = st.text_input("兑换码", placeholder="FMT-XXXX-XXXX", label_visibility="collapsed")
                with col_btn:
                    if st.button("解锁", type="primary", use_container_width=True):
                        if code_input:
                            ok, msg = verify_code(code_input)
                            if ok:
                                st.session_state['unlocked'] = True
                                st.rerun()
                            else:
                                st.error(msg)
                        else:
                            st.warning("请输入兑换码")

            else:
                # ============================
                # 完整报告（已付费）
                # ============================
                st.markdown("✅ **已解锁完整报告**")

                # 筛选器
                col_f1, col_f2, col_f3 = st.columns(3)
                with col_f1:
                    sev_filter = st.selectbox("严重度", ['全部', '错误', '警告', '建议'], key='sev_f')
                with col_f2:
                    mod_filter = st.selectbox("模块", ['全部'] + [m['name'] for m in data['modules']], key='mod_f')
                with col_f3:
                    src_filter = st.selectbox("来源", ['全部', '官方规定', '专业补充', '批注修订'], key='src_f')

                sev_map = {'错误': 'error', '警告': 'warning', '建议': 'info'}
                src_map = {'官方规定': 'official', '专业补充': 'supplement', '批注修订': 'annotation'}

                filtered = issues
                if sev_filter != '全部':
                    filtered = [i for i in filtered if i['severity'] == sev_map[sev_filter]]
                if mod_filter != '全部':
                    filtered = [i for i in filtered if i['module'] == mod_filter]
                if src_filter != '全部':
                    filtered = [i for i in filtered if i['source'] == src_map[src_filter]]

                st.caption(f"显示 {len(filtered)} / {len(issues)} 条")

                for issue in filtered:
                    sev_colors = {'error': '#f87171', 'warning': '#fbbf24', 'info': '#60a5fa'}
                    src_colors = {'official': '#a78bfa', 'supplement': '#2dd4bf', 'annotation': '#fb923c'}
                    border_color = sev_colors.get(issue['severity'], '#64748b')

                    st.markdown(f"""
                    <div style="border-left:3px solid {border_color}; padding:12px 16px; margin-bottom:8px;
                        background:rgba(26,35,50,0.5); border-radius:0 8px 8px 0;">
                        <div style="display:flex; gap:8px; margin-bottom:6px; flex-wrap:wrap;">
                            <span style="background:{border_color}22; color:{border_color}; padding:2px 10px;
                                border-radius:4px; font-size:0.75rem; font-weight:600;">{issue['severity_label']}</span>
                            <span style="background:#33415522; color:#94a3b8; padding:2px 10px;
                                border-radius:4px; font-size:0.75rem;">{issue['module']}</span>
                            <span style="background:{src_colors.get(issue['source'], '#2dd4bf')}22;
                                color:{src_colors.get(issue['source'], '#2dd4bf')}; padding:2px 10px;
                                border-radius:4px; font-size:0.75rem; font-weight:600;">{issue['source_label']}</span>
                            <span style="color:#64748b; font-size:0.75rem; font-family:monospace;">{issue['location']}</span>
                        </div>
                        <div style="font-size:0.9rem; color:#e2e8f0; margin-bottom:4px;">{issue['rule']}</div>
                        <div style="font-size:0.8rem;">
                            <span style="color:#10b981;">期望: {issue['expected']}</span> &nbsp;→&nbsp;
                            <span style="color:#ef4444;">实际: {issue['actual']}</span>
                        </div>
                        {'<div style=\"font-size:0.75rem; color:#64748b; margin-top:4px;\">'+issue["text_preview"]+'</div>' if issue['text_preview'] and issue['text_preview'] != '(空)' else ''}
                    </div>
                    """, unsafe_allow_html=True)

                # 下载按钮
                st.markdown("---")
                st.download_button(
                    "📥 下载完整 HTML 报告",
                    data=html_content,
                    file_name=f"格式审查报告_{report_id}.html",
                    mime="text/html",
                    type="primary",
                    use_container_width=True
                )

        # 页脚
        st.markdown("---")
        st.markdown(f"""
        <div style="text-align:center; color:#64748b; font-size:0.75rem; padding:16px;">
            论文格式智能审查平台 v1.0 &nbsp;|&nbsp; 报告编号: {report_id} &nbsp;|&nbsp;
            生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} &nbsp;|&nbsp;
            检测用时: {elapsed:.1f}s
        </div>
        """, unsafe_allow_html=True)

    finally:
        # 清理临时文件
        try:
            os.unlink(tmp_path)
            if os.path.exists(html_path):
                os.unlink(html_path)
        except:
            pass

else:
    # 未上传时显示介绍
    st.markdown("""
    <div style="text-align:center; padding:60px 20px;">
        <div style="font-size:4rem; margin-bottom:16px;">📄</div>
        <div style="font-size:1.3rem; color:#94a3b8; margin-bottom:24px;">
            拖拽或点击上方上传你的论文 Word 文档
        </div>
        <div style="display:flex; justify-content:center; gap:40px; flex-wrap:wrap;">
            <div style="text-align:center;">
                <div style="font-size:2rem; font-weight:700; color:#3b82f6;">60+</div>
                <div style="color:#64748b; font-size:0.85rem;">检查规则</div>
            </div>
            <div style="text-align:center;">
                <div style="font-size:2rem; font-weight:700; color:#8b5cf6;">13</div>
                <div style="color:#64748b; font-size:0.85rem;">检测模块</div>
            </div>
            <div style="text-align:center;">
                <div style="font-size:2rem; font-weight:700; color:#ec4899;">3秒</div>
                <div style="color:#64748b; font-size:0.85rem;">出报告</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    st.markdown("""
    #### 检测模块覆盖

    | 模块 | 检查内容 |
    |------|----------|
    | 页面设置 | A4纸张、页边距2.5cm |
    | 封面 | 题目字体字号、必填字段 |
    | 摘要 | 中英文摘要格式、关键词 |
    | 目录 | 标题格式、层级 |
    | 正文格式 | 字体字号、行距、首行缩进 |
    | 标题层级 | 一/二/三级标题字号 |
    | 图表规范 | 中英文对照、三线表 |
    | 页眉页脚 | 奇偶页内容、格式 |
    | 参考文献 | 数量、格式、排序 |
    | 结构完整性 | 必需章节、分章检查 |
    | 编号规范 | 图表编号连续、格式一致 |
    | 单位符号 | 国际单位制、化学式 |
    | 内容规范 | 缩写全称、学名斜体、标点 |
    """)
