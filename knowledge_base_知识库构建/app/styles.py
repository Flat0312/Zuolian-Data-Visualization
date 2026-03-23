from __future__ import annotations

import base64
from pathlib import Path

import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"

PRIMARY = "#8f2529"
ACCENT = "#586978"
UMBER = "#78614a"
INK = "#2b241f"
MUTED = "#6c5d4d"
PAPER = "#efe2c8"
PAPER_LIGHT = "#f7f0e2"
PAPER_DARK = "#ddc9a9"
BORDER = "#8d785d"
RULE = "#bda687"
SERIF_STACK = '"Noto Serif SC","Songti SC","SimSun","STSong",serif'
DISPLAY_STACK = '"ZCOOL XiaoWei","Noto Serif SC","Songti SC","SimSun",serif'
CHART_FONT = "Noto Serif SC, Songti SC, SimSun, STSong, serif"


def asset_uri(name: str) -> str:
    path = ASSETS_DIR / name
    if not path.exists():
        return ""
    mime = path.suffix.lower().lstrip(".") or "png"
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/{mime};base64,{data}"


def apply_style() -> None:
    bg = asset_uri("historical_archive_bg.svg")
    paper = asset_uri("paper_texture.png")
    stamp = asset_uri("stamp.png")
    relation_bg = asset_uri("zuolian_relationship_bg.svg")
    st.markdown(
        f"""
        <style>
            :root {{
                --paper: {PAPER};
                --paper-light: {PAPER_LIGHT};
                --paper-dark: {PAPER_DARK};
                --ink: {INK};
                --muted: {MUTED};
                --seal: {PRIMARY};
                --blue: {ACCENT};
                --umber: {UMBER};
                --border: {BORDER};
                --rule: {RULE};
                --serif: {SERIF_STACK};
                --display: {DISPLAY_STACK};
            }}
            html, body, [data-testid="stAppViewContainer"], .stApp {{
                font-family: var(--serif);
                color: var(--ink);
            }}
            .material-symbols-rounded {{
                font-family: "Material Symbols Rounded" !important;
            }}
            .material-symbols-outlined {{
                font-family: "Material Symbols Outlined" !important;
            }}
            .material-symbols-sharp {{
                font-family: "Material Symbols Sharp" !important;
            }}
            .material-icons {{
                font-family: "Material Icons" !important;
            }}
            .material-icons-round {{
                font-family: "Material Icons Round" !important;
            }}
            .material-icons-outlined {{
                font-family: "Material Icons Outlined" !important;
            }}
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
            .stApp {{
                background:
                    linear-gradient(180deg, rgba(247,239,225,.95), rgba(235,221,194,.98)),
                    url("{paper}") center/240px repeat,
                    url("{bg}") center top/cover fixed no-repeat;
                color: var(--ink);
                letter-spacing: .01em;
            }}
            .stApp::before {{
                content: "";
                position: fixed;
                inset: 0;
                background:
                    linear-gradient(90deg, rgba(120,94,62,.08) 0, rgba(120,94,62,.03) 4%, transparent 12%, transparent 88%, rgba(120,94,62,.03) 96%, rgba(120,94,62,.08) 100%),
                    radial-gradient(circle at 15% 18%, rgba(143,37,41,.05), transparent 24%),
                    radial-gradient(circle at 82% 84%, rgba(88,105,120,.06), transparent 24%);
                pointer-events: none;
            }}
            .stApp::after {{
                content: "";
                position: fixed;
                right: 1.5rem;
                bottom: 1rem;
                width: 220px;
                height: 220px;
                background: url("{stamp}") center/contain no-repeat;
                opacity: .08;
                pointer-events: none;
            }}
            header[data-testid="stHeader"] {{
                background: rgba(247,240,228,.55);
                border-bottom: 1px solid rgba(111,90,64,.18);
                backdrop-filter: blur(2px);
            }}
            [data-testid="block-container"] {{
                position: relative;
                max-width: 1360px;
                padding-top: 2rem;
                padding-bottom: 3rem;
                z-index: 1;
            }}
            [data-testid="block-container"]::before {{
                content: none;
            }}
            h1, h2, h3, h4 {{
                font-family: var(--display);
                font-weight: 400;
                letter-spacing: .08em;
                color: #652327;
            }}
            h1 {{
                font-size: clamp(2.35rem, 3vw, 3.35rem);
                line-height: 1.14;
            }}
            h2 {{
                font-size: 1.6rem;
                line-height: 1.3;
            }}
            h3 {{
                font-size: 1.2rem;
                line-height: 1.45;
            }}
            p, label, .stMarkdown, .stCaption {{
                line-height: 1.85;
            }}
            .hero {{
                position: relative;
                padding: 1.8rem 1.7rem 1.5rem 1.7rem;
                margin-bottom: 1.2rem;
                border: 1px solid rgba(111,90,64,.45);
                background:
                    linear-gradient(180deg, rgba(249,244,233,.96), rgba(240,229,210,.92)),
                    url("{paper}") center/220px repeat,
                    url("{relation_bg}") right 1.2rem top .7rem / 220px no-repeat;
                box-shadow:
                    inset 0 0 0 1px rgba(255,249,239,.82),
                    inset 0 0 26px rgba(124,95,61,.08);
            }}
            .hero::before {{
                content: "";
                position: absolute;
                top: .95rem;
                right: 1rem;
                width: 120px;
                height: 120px;
                background: url("{stamp}") center/contain no-repeat;
                opacity: .06;
                pointer-events: none;
            }}
            .hero small, .muted {{
                color: var(--muted);
                line-height: 1.8;
            }}
            .hero small {{
                display: inline-block;
                margin-bottom: .35rem;
                letter-spacing: .18em;
                text-transform: uppercase;
            }}
            .hero h1 {{
                margin: .3rem 0 .55rem 0;
            }}
            .page-note {{
                padding: .95rem 1rem;
                margin-bottom: 1rem;
                border: 1px solid rgba(111,90,64,.4);
                border-left: 4px solid var(--seal);
                background: rgba(245,237,222,.86);
                box-shadow: inset 0 0 0 1px rgba(255,249,239,.7);
                color: var(--muted);
            }}
            .route-card {{
                height: 100%;
                padding: 1rem 1.05rem;
                margin-bottom: .9rem;
                border: 1px solid rgba(111,90,64,.38);
                background:
                    linear-gradient(180deg, rgba(249,243,232,.96), rgba(239,228,209,.92)),
                    url("{paper}") center/180px repeat;
                box-shadow:
                    inset 0 0 0 1px rgba(255,248,236,.78),
                    inset 0 0 18px rgba(111,90,64,.05);
            }}
            .route-card-kicker {{
                color: var(--muted);
                font-size: .85rem;
                letter-spacing: .14em;
                text-transform: uppercase;
            }}
            .route-card-title {{
                margin: .45rem 0;
                font-family: var(--display);
                font-size: 1.18rem;
                color: var(--seal);
                letter-spacing: .06em;
            }}
            .route-card-body {{
                color: var(--ink);
                line-height: 1.8;
            }}
            .section-card {{
                padding: 1rem 1.05rem;
                margin-bottom: 1rem;
                border: 1px solid rgba(111,90,64,.38);
                background: rgba(246,239,227,.84);
                box-shadow: inset 0 0 0 1px rgba(255,248,236,.72);
            }}
            .section-card-title {{
                margin-bottom: .45rem;
                font-family: var(--display);
                font-size: 1.12rem;
                color: var(--seal);
                letter-spacing: .06em;
            }}
            .tag-row {{
                display: flex;
                flex-wrap: wrap;
                gap: .4rem;
                margin-top: .65rem;
            }}
            .tag-item {{
                display: inline-flex;
                align-items: center;
                padding: .18rem .48rem;
                border: 1px solid rgba(111,90,64,.34);
                background: rgba(247,240,228,.82);
                color: var(--muted);
                font-size: .86rem;
                line-height: 1.5;
            }}
            section[data-testid="stSidebar"] {{
                background:
                    linear-gradient(180deg, rgba(66,51,38,.98), rgba(37,29,23,.99)),
                    url("{paper}") center/220px repeat;
                border-right: 1px solid rgba(236,216,183,.12);
            }}
            section[data-testid="stSidebar"] .block-container {{
                position: relative;
                padding-top: 2rem;
            }}
            section[data-testid="stSidebar"] .block-container::before {{
                content: "卷宗目录";
                display: block;
                margin-bottom: 1rem;
                padding-bottom: .65rem;
                border-bottom: 1px solid rgba(236,216,183,.22);
                font-family: var(--display);
                font-size: 1.15rem;
                letter-spacing: .18em;
                color: #f2ddba;
            }}
            section[data-testid="stSidebar"] * {{
                color: #f7eddc;
            }}
            section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
            section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {{
                color: #f4debb;
            }}
            section[data-testid="stSidebar"] div[role="radiogroup"] label {{
                margin-bottom: .4rem;
                padding: .55rem .75rem .55rem .95rem;
                border: 1px solid rgba(236,216,183,.16);
                border-left: 3px solid transparent;
                background: rgba(248,233,205,.05);
            }}
            section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {{
                background: rgba(248,233,205,.1);
            }}
            section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {{
                border-left-color: #d8b989;
                background: rgba(248,233,205,.12);
            }}
            section[data-testid="stSidebar"] .stCaption {{
                color: #dbc7a7;
            }}
            .stTabs [data-baseweb="tab-list"] {{
                gap: .3rem;
                padding-left: .25rem;
            }}
            .stTabs [data-baseweb="tab"] {{
                height: auto;
                padding: .65rem 1rem .58rem;
                margin-top: .35rem;
                border: 1px solid rgba(111,90,64,.42);
                border-bottom: none;
                border-radius: 0;
                background: linear-gradient(180deg, rgba(232,218,194,.94), rgba(220,202,171,.92));
                box-shadow: inset 0 1px 0 rgba(255,248,237,.65);
                color: var(--muted);
            }}
            .stTabs [data-baseweb="tab"][aria-selected="true"] {{
                transform: translateY(-2px);
                background: linear-gradient(180deg, rgba(248,241,228,.98), rgba(241,230,210,.96));
                color: var(--seal);
            }}
            .stTabs [data-baseweb="tab-panel"] {{
                padding: 1rem 1.1rem 1.2rem;
                margin-top: -1px;
                border: 1px solid rgba(111,90,64,.42);
                background:
                    linear-gradient(180deg, rgba(250,245,235,.92), rgba(242,231,212,.9)),
                    url("{paper}") center/220px repeat;
                box-shadow: inset 0 0 0 1px rgba(255,249,239,.72);
            }}
            div[data-testid="stMetric"] {{
                min-height: 7.1rem;
                padding: .9rem 1rem 1rem;
                border: 1px solid rgba(111,90,64,.42);
                border-radius: 0;
                background:
                    linear-gradient(180deg, rgba(248,241,230,.96), rgba(236,222,197,.9)),
                    url("{paper}") center/180px repeat;
                box-shadow:
                    inset 0 0 0 1px rgba(255,248,236,.8),
                    inset 0 0 18px rgba(111,90,64,.05);
            }}
            div[data-testid="stMetricLabel"] {{
                color: var(--muted);
                font-size: .92rem;
                letter-spacing: .08em;
            }}
            div[data-testid="stMetricValue"] {{
                font-family: var(--display);
                color: var(--seal);
                line-height: 1.2;
            }}
            div[data-testid="stDataFrame"],
            div[data-testid="stTable"],
            div[data-testid="stVegaLiteChart"],
            div[data-testid="stPlotlyChart"],
            div[data-testid="stIFrame"] {{
                padding: .55rem;
                border: 1px solid rgba(111,90,64,.42);
                background:
                    linear-gradient(180deg, rgba(248,241,231,.96), rgba(239,228,209,.9)),
                    url("{paper}") center/200px repeat;
                box-shadow:
                    inset 0 0 0 1px rgba(255,248,236,.8),
                    inset 0 0 18px rgba(111,90,64,.05);
            }}
            div[data-testid="stDataFrame"] * {{
                font-family: var(--serif) !important;
            }}
            iframe {{
                border: none !important;
                background: transparent !important;
            }}
            div[data-testid="stTextInput"] > label,
            div[data-testid="stSelectbox"] > label,
            div[data-testid="stSlider"] > label {{
                font-size: .95rem;
                font-weight: 600;
                color: var(--ink);
                letter-spacing: .06em;
            }}
            div[data-testid="stTextInput"] input,
            div[data-testid="stNumberInput"] input,
            div[data-testid="stTextArea"] textarea,
            div[data-baseweb="select"] > div,
            div[data-testid="stDateInput"] input {{
                border-radius: 0 !important;
                border: 1px solid rgba(111,90,64,.38) !important;
                background: rgba(247,240,228,.95) !important;
                color: var(--ink) !important;
                box-shadow: inset 0 0 0 1px rgba(255,248,236,.75);
            }}
            div[data-baseweb="select"] * {{
                font-family: var(--serif) !important;
            }}
            .stSlider [data-baseweb="slider"] > div div {{
                background: var(--seal);
            }}
            .stSlider [role="slider"] {{
                border: 1px solid rgba(91,67,49,.4);
                background: var(--paper-light);
            }}
            .stButton > button {{
                border-radius: 0;
                border: 1px solid rgba(111,90,64,.48);
                background: linear-gradient(180deg, rgba(241,231,210,.96), rgba(227,211,182,.92));
                color: var(--ink);
                box-shadow: inset 0 1px 0 rgba(255,248,236,.78);
            }}
            .stButton > button:hover {{
                color: var(--seal);
                border-color: var(--seal);
            }}
            div[data-testid="stAlert"] {{
                border-radius: 0;
                border: 1px solid rgba(111,90,64,.42);
                background: rgba(245,237,223,.9);
                box-shadow: inset 0 0 0 1px rgba(255,248,236,.76);
            }}
            .relation-entry {{
                padding: .8rem .9rem;
                margin-bottom: .6rem;
                border: 1px solid rgba(111,90,64,.36);
                background: rgba(246,239,227,.84);
                box-shadow: inset 0 0 0 1px rgba(255,248,236,.72);
            }}
            .relation-entry-title {{
                margin-bottom: .22rem;
                font-family: var(--display);
                font-size: 1.08rem;
                color: var(--ink);
                letter-spacing: .04em;
            }}
            .relation-entry-meta {{
                color: var(--muted);
                font-size: .93rem;
                line-height: 1.7;
            }}
            .relation-detail-head {{
                padding: 1rem 1.1rem .9rem;
                margin: .6rem 0 1rem;
                border: 1px solid rgba(111,90,64,.44);
                background:
                    linear-gradient(180deg, rgba(249,243,232,.96), rgba(239,228,209,.92)),
                    url("{paper}") center/200px repeat;
                box-shadow: inset 0 0 0 1px rgba(255,248,236,.78);
            }}
            .relation-detail-title {{
                font-family: var(--display);
                font-size: 1.42rem;
                color: var(--seal);
                letter-spacing: .08em;
            }}
            .relation-detail-subtitle {{
                margin-top: .3rem;
                color: var(--muted);
                line-height: 1.75;
            }}
            .provenance-note {{
                margin-top: .45rem;
                color: var(--muted);
                font-size: .88rem;
            }}
            .evidence-summary {{
                padding: .85rem 1rem;
                margin-bottom: .75rem;
                border-left: 3px solid var(--seal);
                background: rgba(246,239,227,.82);
            }}
            .evidence-meta {{
                color: var(--muted);
                line-height: 1.8;
                font-size: .95rem;
            }}
            .excerpt-block {{
                margin-top: .65rem;
                padding: .8rem .9rem;
                border: 1px solid rgba(111,90,64,.32);
                background: rgba(252,248,240,.76);
                color: var(--ink);
                line-height: 1.85;
                max-height: 15rem;
                overflow: auto;
            }}
            .analysis-note {{
                padding: 1rem 1.05rem;
                border: 1px solid rgba(111,90,64,.42);
                border-left: 4px solid var(--blue);
                background: rgba(246,239,227,.86);
                box-shadow: inset 0 0 0 1px rgba(255,248,236,.72);
            }}
            .analysis-note-title {{
                margin-bottom: .35rem;
                font-family: var(--display);
                font-size: 1.08rem;
                letter-spacing: .06em;
                color: var(--seal);
            }}
            .analysis-note-body {{
                color: var(--ink);
                line-height: 1.82;
            }}
            .analysis-note-source {{
                margin-top: .55rem;
                color: var(--muted);
                font-size: .9rem;
                line-height: 1.7;
            }}
            .finding-card {{
                height: 100%;
                padding: 1rem 1.05rem;
                margin-bottom: 1rem;
                border: 1px solid rgba(111,90,64,.42);
                background:
                    linear-gradient(180deg, rgba(249,243,232,.96), rgba(239,228,209,.92)),
                    url("{paper}") center/180px repeat;
                box-shadow:
                    inset 0 0 0 1px rgba(255,248,236,.78),
                    inset 0 0 18px rgba(111,90,64,.05);
            }}
            .finding-card-title {{
                margin-bottom: .5rem;
                font-family: var(--display);
                font-size: 1.18rem;
                color: var(--seal);
                letter-spacing: .06em;
            }}
            .finding-card-body {{
                color: var(--ink);
                line-height: 1.82;
            }}
            .finding-card-label {{
                margin-top: .7rem;
                color: var(--muted);
                font-size: .9rem;
                letter-spacing: .04em;
            }}
            @media (max-width: 900px) {{
                [data-testid="block-container"] {{
                    padding-top: 1.25rem;
                }}
                [data-testid="block-container"]::before {{
                    content: none;
                }}
                .hero {{
                    padding: 1.35rem 1.15rem;
                    background-size: auto, 200px, 150px;
                }}
                .stTabs [data-baseweb="tab"] {{
                    padding-left: .75rem;
                    padding-right: .75rem;
                }}
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )
