"""Notion × Claude 混合视觉设计系统 v2。

设计语言：
  - 层次感：多级阴影 + 暖色渐变背景 + 微妙纹理
  - 质感：卡片悬浮动效、粗细对比、留白节奏
  - Claude 赭石：大胆使用主色做渐变和强调
"""

from __future__ import annotations

import base64
from pathlib import Path

import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"

# ── 设计令牌 ──────────────────────────────────────────────────
PRIMARY = "#C87941"
PRIMARY_LIGHT = "#E8A96D"
PRIMARY_DARK = "#A85E2A"
PRIMARY_BG = "rgba(200,121,65,0.06)"
PRIMARY_BG_WARM = "rgba(200,121,65,0.03)"

TERRACOTTA = "#B85C38"
SAGE = "#6B8F71"
SLATE_BLUE = "#5E7184"

BG = "#FAFAF8"
BG_CARD = "#FFFFFF"
BG_ELEVATED = "#FFFFFF"
BG_SIDEBAR = "#F5F4F0"
BORDER = "#E8E5DE"
BORDER_LIGHT = "#F0EDE6"
TEXT = "#1A1A1A"
TEXT_SECONDARY = "#6B6B6B"
TEXT_TERTIARY = "#9B9B9B"

SERIF = '"Noto Serif SC","Source Han Serif SC","Songti SC","SimSun",serif'
SANS = '"Inter","SF Pro Display",-apple-system,"Segoe UI",sans-serif'
MONO = '"JetBrains Mono","Fira Code","SF Mono",monospace'

RADIUS_SM = "4px"
RADIUS_MD = "8px"
RADIUS_LG = "12px"

# 向后兼容别名
ACCENT = SLATE_BLUE
UMBER = "#78614a"
INK = TEXT
MUTED = TEXT_SECONDARY
PAPER = BG
PAPER_LIGHT = BG_CARD
PAPER_DARK = BORDER
BORDER_OLD = BORDER
RULE = BORDER_LIGHT
CHART_FONT = "Noto Serif SC, Songti SC, SimSun, STSong, serif"
SERIF_STACK = SERIF
DISPLAY_STACK = SERIF


def asset_uri(name: str) -> str:
    path = ASSETS_DIR / name
    if not path.exists():
        return ""
    mime = path.suffix.lower().lstrip(".") or "png"
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/{mime};base64,{data}"


def apply_style() -> None:
    """注入全局 CSS — 高级感版本。"""
    st.markdown(
        f"""
        <style>
        /* ── CSS 变量 ── */
        :root {{
            --primary: {PRIMARY};
            --primary-light: {PRIMARY_LIGHT};
            --primary-dark: {PRIMARY_DARK};
            --primary-bg: {PRIMARY_BG};
            --primary-bg-warm: {PRIMARY_BG_WARM};
            --terracotta: {TERRACOTTA};
            --sage: {SAGE};
            --slate-blue: {SLATE_BLUE};
            --bg: {BG};
            --bg-card: {BG_CARD};
            --bg-elevated: {BG_ELEVATED};
            --bg-sidebar: {BG_SIDEBAR};
            --border: {BORDER};
            --border-light: {BORDER_LIGHT};
            --text: {TEXT};
            --text-secondary: {TEXT_SECONDARY};
            --text-tertiary: {TEXT_TERTIARY};
            --serif: {SERIF};
            --sans: {SANS};
            --mono: {MONO};
            --radius-sm: {RADIUS_SM};
            --radius-md: {RADIUS_MD};
            --radius-lg: {RADIUS_LG};
        }}

        /* ── 全局基础 ── */
        html, body, [data-testid="stAppViewContainer"], .stApp {{
            font-family: var(--serif);
            color: var(--text);
            -webkit-font-smoothing: antialiased;
        }}

        /* ── 页面背景：暖色渐变 + 微妙噪点纹理 ── */
        .stApp {{
            background:
                linear-gradient(175deg, #FDFCFA 0%, #FAF8F3 35%, #F6F2EA 70%, #F3EDE2 100%);
            letter-spacing: .01em;
        }}

        .stApp::before {{
            content: "";
            position: fixed;
            inset: 0;
            background:
                radial-gradient(ellipse at 0% 0%, rgba(200,121,65,0.03) 0%, transparent 50%),
                radial-gradient(ellipse at 100% 100%, rgba(94,113,132,0.025) 0%, transparent 50%);
            pointer-events: none;
            z-index: 0;
        }}

        /* ── 顶部栏 ── */
        header[data-testid="stHeader"] {{
            background: rgba(253,252,250,0.82);
            border-bottom: 1px solid rgba(232,229,222,0.6);
            backdrop-filter: blur(16px) saturate(1.2);
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        }}

        /* ── 主内容区 ── */
        [data-testid="block-container"] {{
            position: relative;
            max-width: 1280px;
            padding: 2.5rem 2.5rem 3.5rem;
            z-index: 1;
        }}

        /* ── 排版 ── */
        h1, h2, h3, h4 {{
            font-family: var(--serif);
            font-weight: 600;
            color: var(--text);
            letter-spacing: .03em;
            line-height: 1.3;
        }}

        h1 {{
            font-size: clamp(2.2rem, 3vw, 3rem);
            line-height: 1.12;
            font-weight: 700;
            letter-spacing: .01em;
        }}

        h2 {{
            font-size: 1.55rem;
            font-weight: 700;
        }}

        h3 {{
            font-size: 1.15rem;
        }}

        p, label, .stMarkdown, .stCaption {{
            line-height: 1.85;
        }}

        p {{ margin: 0; }}

        /* ── Hero：暖色渐变 + 左侧赭石装饰条 ── */
        .hero {{
            position: relative;
            padding: 2.4rem 2.4rem 2rem;
            margin-bottom: 2rem;
            border: 1px solid rgba(200,121,65,0.12);
            border-radius: var(--radius-lg);
            background:
                linear-gradient(135deg, #FFFCF8 0%, #FFF9F2 40%, #FFF6EC 100%);
            box-shadow:
                0 1px 2px rgba(200,121,65,0.04),
                0 4px 16px rgba(200,121,65,0.06),
                0 12px 40px rgba(0,0,0,0.03);
            overflow: hidden;
        }}

        .hero::before {{
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 4px;
            background: linear-gradient(180deg, var(--primary), var(--terracotta));
            border-radius: var(--radius-lg) 0 0 var(--radius-lg);
        }}

        .hero::after {{
            content: "";
            position: absolute;
            top: -40%;
            right: -10%;
            width: 320px;
            height: 320px;
            background: radial-gradient(circle, rgba(200,121,65,0.04) 0%, transparent 70%);
            pointer-events: none;
        }}

        .hero small, .muted {{
            color: var(--text-secondary);
            line-height: 1.8;
        }}

        .hero small {{
            display: inline-block;
            margin-bottom: .5rem;
            padding: .2rem .65rem;
            background: linear-gradient(135deg, rgba(200,121,65,0.08), rgba(200,121,65,0.04));
            border: 1px solid rgba(200,121,65,0.12);
            border-radius: 3px;
            font-family: var(--sans);
            font-size: .72rem;
            font-weight: 600;
            letter-spacing: .18em;
            text-transform: uppercase;
            color: var(--primary);
        }}

        .hero h1 {{
            margin: .3rem 0 .7rem;
            color: var(--text);
            background: linear-gradient(135deg, #1A1A1A 60%, #3D2E1E 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }}

        /* ── 页面说明条 ── */
        .page-note {{
            padding: 1rem 1.2rem;
            margin-bottom: 1.4rem;
            border: 1px solid rgba(200,121,65,0.1);
            border-left: 3px solid var(--primary);
            border-radius: 0 var(--radius-md) var(--radius-md) 0;
            background: linear-gradient(90deg, rgba(200,121,65,0.04), transparent);
            color: var(--text-secondary);
            font-size: .92rem;
            line-height: 1.75;
        }}

        /* ── 路由卡片：悬浮 + 渐变边框 ── */
        .route-card {{
            position: relative;
            height: 100%;
            padding: 1.4rem 1.5rem;
            margin-bottom: .8rem;
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            background: var(--bg-elevated);
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
            transition: all .25s cubic-bezier(.4,0,.2,1);
        }}

        .route-card:hover {{
            border-color: rgba(200,121,65,0.25);
            box-shadow:
                0 4px 12px rgba(200,121,65,0.08),
                0 12px 32px rgba(0,0,0,0.04);
            transform: translateY(-2px);
        }}

        .route-card-kicker {{
            font-family: var(--sans);
            font-size: .72rem;
            font-weight: 600;
            letter-spacing: .16em;
            text-transform: uppercase;
            color: var(--primary);
        }}

        .route-card-title {{
            margin: .45rem 0;
            font-family: var(--serif);
            font-size: 1.15rem;
            font-weight: 700;
            color: var(--text);
            letter-spacing: .02em;
        }}

        .route-card-body {{
            color: var(--text-secondary);
            line-height: 1.75;
            font-size: .92rem;
        }}

        /* ── 通用卡片 ── */
        .section-card {{
            padding: 1.2rem 1.3rem;
            margin-bottom: 1rem;
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            background: var(--bg-elevated);
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
            transition: box-shadow .2s ease;
        }}

        .section-card:hover {{
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }}

        .section-card-title {{
            margin-bottom: .4rem;
            font-family: var(--serif);
            font-size: 1.08rem;
            font-weight: 600;
            color: var(--text);
            letter-spacing: .02em;
        }}

        /* ── 标签 ── */
        .tag-row {{
            display: flex;
            flex-wrap: wrap;
            gap: .4rem;
            margin-top: .65rem;
        }}

        .tag-item {{
            display: inline-flex;
            align-items: center;
            padding: .2rem .55rem;
            border: 1px solid var(--border);
            border-radius: 3px;
            background: linear-gradient(180deg, #FAFAF8, #F7F6F3);
            font-family: var(--sans);
            font-size: .8rem;
            font-weight: 500;
            color: var(--text-secondary);
            line-height: 1.5;
            transition: all .15s ease;
        }}

        .tag-item:hover {{
            border-color: var(--primary-light);
            color: var(--primary);
            background: rgba(200,121,65,0.04);
        }}

        /* ── 侧边栏：深色暖褐渐变 + 赭石高光 ── */
        section[data-testid="stSidebar"] {{
            background: linear-gradient(195deg, #2A2320 0%, #1E1A17 60%, #171412 100%);
            border-right: 1px solid rgba(200,121,65,0.08);
        }}

        section[data-testid="stSidebar"] .block-container {{
            position: relative;
            padding-top: 1.8rem;
        }}

        section[data-testid="stSidebar"] .block-container::before {{
            content: "";
            display: block;
            height: 3px;
            margin-bottom: 1.2rem;
            border-radius: 2px;
            background: linear-gradient(90deg, var(--primary), var(--primary-light), transparent);
        }}

        section[data-testid="stSidebar"] * {{
            color: #E8DDD0;
        }}

        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {{
            color: #F2E8DA;
            font-weight: 700;
            letter-spacing: .06em;
        }}

        section[data-testid="stSidebar"] div[role="radiogroup"] label {{
            margin-bottom: .25rem;
            padding: .55rem .75rem .55rem 1rem;
            border: 1px solid transparent;
            border-left: 3px solid transparent;
            border-radius: var(--radius-sm);
            background: rgba(255,255,255,0.02);
            font-family: var(--sans);
            font-size: .88rem;
            font-weight: 400;
            transition: all .2s ease;
        }}

        section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {{
            background: rgba(200,121,65,0.06);
            border-left-color: rgba(200,121,65,0.2);
        }}

        section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {{
            border-left-color: var(--primary);
            background: linear-gradient(90deg, rgba(200,121,65,0.12), rgba(200,121,65,0.03));
            color: #F5E6D3;
            font-weight: 500;
        }}

        section[data-testid="stSidebar"] .stCaption {{
            color: #8A7E72;
        }}

        /* ── 标签页 ── */
        .stTabs [data-baseweb="tab-list"] {{
            gap: 0;
            padding-left: 0;
            border-bottom: 1px solid var(--border);
        }}

        .stTabs [data-baseweb="tab"] {{
            position: relative;
            height: auto;
            padding: .65rem 1.1rem .6rem;
            margin-bottom: -1px;
            border: none;
            border-bottom: 2px solid transparent;
            border-radius: 0;
            background: transparent;
            font-family: var(--sans);
            font-size: .88rem;
            font-weight: 400;
            color: var(--text-secondary);
            transition: color .2s ease, border-color .2s ease;
        }}

        .stTabs [data-baseweb="tab"]:hover {{
            color: var(--text);
        }}

        .stTabs [data-baseweb="tab"][aria-selected="true"] {{
            border-bottom-color: var(--primary);
            background: transparent;
            color: var(--primary);
            font-weight: 600;
        }}

        .stTabs [data-baseweb="tab-panel"] {{
            padding: 1.3rem 0 0;
            border: none;
            background: transparent;
        }}

        /* ── 指标卡片：顶部赭石渐变条 ── */
        div[data-testid="stMetric"] {{
            position: relative;
            min-height: 7rem;
            padding: 1.1rem 1.2rem 1rem;
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            background: var(--bg-elevated);
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
            overflow: hidden;
            transition: box-shadow .2s ease, transform .2s ease;
        }}

        div[data-testid="stMetric"]:hover {{
            box-shadow: 0 4px 16px rgba(200,121,65,0.08);
            transform: translateY(-1px);
        }}

        div[data-testid="stMetric"]::before {{
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg, var(--primary), var(--primary-light));
        }}

        div[data-testid="stMetricLabel"] {{
            font-family: var(--sans);
            font-size: .72rem;
            font-weight: 600;
            letter-spacing: .1em;
            text-transform: uppercase;
            color: var(--text-tertiary);
        }}

        div[data-testid="stMetricValue"] {{
            font-family: var(--serif);
            font-size: 1.8rem;
            font-weight: 700;
            color: var(--primary);
            line-height: 1.2;
        }}

        /* ── 数据展示 ── */
        div[data-testid="stDataFrame"],
        div[data-testid="stTable"],
        div[data-testid="stVegaLiteChart"],
        div[data-testid="stPlotlyChart"],
        div[data-testid="stIFrame"] {{
            padding: .6rem;
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            background: var(--bg-elevated);
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        }}

        div[data-testid="stDataFrame"] * {{
            font-family: var(--sans) !important;
        }}

        iframe {{
            border: none !important;
            background: transparent !important;
        }}

        /* ── 表单控件 ── */
        div[data-testid="stTextInput"] > label,
        div[data-testid="stSelectbox"] > label,
        div[data-testid="stSlider"] > label {{
            font-family: var(--sans);
            font-size: .82rem;
            font-weight: 600;
            color: var(--text-secondary);
            letter-spacing: .04em;
        }}

        div[data-testid="stTextInput"] input,
        div[data-testid="stNumberInput"] input,
        div[data-testid="stTextArea"] textarea,
        div[data-baseweb="select"] > div,
        div[data-testid="stDateInput"] input {{
            border-radius: var(--radius-md) !important;
            border: 1px solid var(--border) !important;
            background: var(--bg-elevated) !important;
            color: var(--text) !important;
            font-family: var(--serif) !important;
            box-shadow: 0 1px 2px rgba(0,0,0,0.02);
            transition: all .2s ease !important;
        }}

        div[data-testid="stTextInput"] input:focus,
        div[data-testid="stNumberInput"] input:focus,
        div[data-testid="stTextArea"] textarea:focus {{
            border-color: var(--primary) !important;
            box-shadow:
                0 0 0 3px rgba(200,121,65,0.08),
                0 1px 2px rgba(0,0,0,0.02) !important;
        }}

        div[data-baseweb="select"] * {{
            font-family: var(--serif) !important;
        }}

        .stSlider [data-baseweb="slider"] > div div {{
            background: linear-gradient(90deg, var(--primary), var(--primary-light));
        }}

        .stSlider [role="slider"] {{
            border: 2px solid var(--primary);
            background: var(--bg-elevated);
            box-shadow: 0 1px 4px rgba(200,121,65,0.2);
        }}

        /* ── 按钮 ── */
        .stButton > button {{
            border-radius: var(--radius-md);
            border: 1px solid var(--border);
            background: linear-gradient(180deg, #FFFFFF, #FAFAF8);
            font-family: var(--sans);
            font-size: .88rem;
            font-weight: 500;
            color: var(--text);
            box-shadow: 0 1px 2px rgba(0,0,0,0.03);
            transition: all .2s cubic-bezier(.4,0,.2,1);
        }}

        .stButton > button:hover {{
            border-color: var(--primary);
            color: var(--primary);
            box-shadow: 0 2px 8px rgba(200,121,65,0.12);
            transform: translateY(-1px);
        }}

        /* ── 提示框 ── */
        div[data-testid="stAlert"] {{
            border-radius: var(--radius-md);
            border: 1px solid var(--border);
            background: linear-gradient(135deg, #FAFAF8, #F7F6F3);
            box-shadow: 0 1px 2px rgba(0,0,0,0.02);
        }}

        /* ── 关系条目 ── */
        .relation-entry {{
            position: relative;
            padding: 1rem 1.1rem;
            margin-bottom: .55rem;
            border: 1px solid var(--border);
            border-radius: var(--radius-md);
            background: var(--bg-elevated);
            box-shadow: 0 1px 2px rgba(0,0,0,0.015);
            transition: all .2s ease;
        }}

        .relation-entry:hover {{
            border-color: rgba(200,121,65,0.2);
            box-shadow: 0 3px 12px rgba(200,121,65,0.06);
            transform: translateY(-1px);
        }}

        .relation-entry-title {{
            margin-bottom: .22rem;
            font-family: var(--serif);
            font-size: 1.02rem;
            font-weight: 600;
            color: var(--text);
        }}

        .relation-entry-meta {{
            font-family: var(--sans);
            font-size: .85rem;
            color: var(--text-secondary);
            line-height: 1.7;
        }}

        .relation-detail-head {{
            position: relative;
            padding: 1.3rem 1.4rem 1.1rem;
            margin: .5rem 0 1.2rem;
            border: 1px solid rgba(200,121,65,0.12);
            border-radius: var(--radius-lg);
            background: linear-gradient(135deg, #FFFCF8, #FFF9F2);
            box-shadow:
                0 2px 8px rgba(200,121,65,0.05),
                0 8px 24px rgba(0,0,0,0.02);
            overflow: hidden;
        }}

        .relation-detail-head::before {{
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 3px;
            background: linear-gradient(180deg, var(--primary), var(--primary-light));
        }}

        .relation-detail-title {{
            font-family: var(--serif);
            font-size: 1.4rem;
            font-weight: 700;
            color: var(--text);
        }}

        .relation-detail-subtitle {{
            margin-top: .35rem;
            font-family: var(--sans);
            color: var(--text-secondary);
            line-height: 1.7;
        }}

        .provenance-note {{
            margin-top: .45rem;
            font-family: var(--sans);
            font-size: .82rem;
            color: var(--text-tertiary);
        }}

        .evidence-summary {{
            padding: .9rem 1.1rem;
            margin-bottom: .7rem;
            border-left: 3px solid var(--primary);
            border-radius: 0 var(--radius-md) var(--radius-md) 0;
            background: linear-gradient(90deg, rgba(200,121,65,0.05), transparent);
        }}

        .evidence-meta {{
            font-family: var(--sans);
            font-size: .85rem;
            color: var(--text-secondary);
            line-height: 1.75;
        }}

        .excerpt-block {{
            margin-top: .6rem;
            padding: .9rem 1.1rem;
            border: 1px solid var(--border-light);
            border-left: 3px solid rgba(200,121,65,0.15);
            border-radius: 0 var(--radius-md) var(--radius-md) 0;
            background: linear-gradient(135deg, #FDFCFA, #FAF8F3);
            color: var(--text);
            line-height: 1.9;
            max-height: 15rem;
            overflow: auto;
        }}

        /* ── 分析笔记 ── */
        .analysis-note {{
            padding: 1.1rem 1.2rem;
            border: 1px solid var(--border);
            border-left: 3px solid var(--slate-blue);
            border-radius: 0 var(--radius-md) var(--radius-md) 0;
            background: linear-gradient(90deg, rgba(94,113,132,0.03), var(--bg-elevated));
            box-shadow: 0 1px 2px rgba(0,0,0,0.015);
        }}

        .analysis-note-title {{
            margin-bottom: .35rem;
            font-family: var(--sans);
            font-size: .75rem;
            font-weight: 600;
            letter-spacing: .12em;
            text-transform: uppercase;
            color: var(--slate-blue);
        }}

        .analysis-note-body {{
            color: var(--text);
            line-height: 1.82;
        }}

        .analysis-note-source {{
            margin-top: .55rem;
            font-family: var(--sans);
            font-size: .82rem;
            color: var(--text-tertiary);
            line-height: 1.7;
        }}

        /* ── 研究发现卡片 ── */
        .finding-card {{
            height: 100%;
            padding: 1.3rem 1.4rem;
            margin-bottom: 1rem;
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            background: var(--bg-elevated);
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
            transition: all .25s cubic-bezier(.4,0,.2,1);
        }}

        .finding-card:hover {{
            border-color: rgba(200,121,65,0.2);
            box-shadow:
                0 4px 12px rgba(200,121,65,0.06),
                0 12px 32px rgba(0,0,0,0.03);
            transform: translateY(-2px);
        }}

        .finding-card-title {{
            margin-bottom: .5rem;
            font-family: var(--serif);
            font-size: 1.12rem;
            font-weight: 700;
            color: var(--text);
        }}

        .finding-card-body {{
            color: var(--text-secondary);
            line-height: 1.82;
        }}

        .finding-card-label {{
            margin-top: .7rem;
            font-family: var(--sans);
            font-size: .75rem;
            font-weight: 600;
            letter-spacing: .08em;
            text-transform: uppercase;
            color: var(--text-tertiary);
        }}

        /* ── 装饰性分隔线 ── */
        hr, [data-testid="stDivider"] {{
            border: none;
            height: 1px;
            background: linear-gradient(90deg, transparent, var(--border), transparent);
            margin: 1.5rem 0;
        }}

        /* ── 响应式 ── */
        @media (max-width: 900px) {{
            [data-testid="block-container"] {{
                padding: 1.2rem 1rem 2rem;
            }}
            .hero {{
                padding: 1.5rem 1.2rem;
            }}
            .stTabs [data-baseweb="tab"] {{
                padding-left: .7rem;
                padding-right: .7rem;
                font-size: .82rem;
            }}
        }}

        /* ── Material Icons 基线 ── */
        .material-symbols-rounded,
        .material-symbols-outlined,
        .material-symbols-sharp,
        .material-icons,
        .material-icons-round,
        .material-icons-outlined {{
            font-style: normal !important;
            font-weight: 400 !important;
            letter-spacing: normal !important;
            text-transform: none !important;
            white-space: nowrap !important;
            word-wrap: normal !important;
            direction: ltr !important;
            line-height: 1 !important;
            -webkit-font-smoothing: antialiased;
        }}

        .material-symbols-rounded {{ font-family: "Material Symbols Rounded" !important; }}
        .material-symbols-outlined {{ font-family: "Material Symbols Outlined" !important; }}
        .material-symbols-sharp {{ font-family: "Material Symbols Sharp" !important; }}
        .material-icons {{ font-family: "Material Icons" !important; }}
        .material-icons-round {{ font-family: "Material Icons Round" !important; }}
        .material-icons-outlined {{ font-family: "Material Icons Outlined" !important; }}
        </style>
        """,
        unsafe_allow_html=True,
    )
