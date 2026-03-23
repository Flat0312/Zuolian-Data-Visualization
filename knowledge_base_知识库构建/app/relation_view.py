from __future__ import annotations

import math
from dataclasses import dataclass, field
from html import escape
from pathlib import Path

import networkx as nx
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from pyvis.network import Network

from data_loader import clean_text, match, show, split_ids
from event_view import render_event_detail, render_event_map
from historical_map import HistoricalEvent
from relation_evidence import RelationDetail, build_relation_detail_index, canonical_pair_key
from styles import ACCENT, BORDER, CHART_FONT, INK, MUTED, PAPER, PAPER_LIGHT, PRIMARY, UMBER, asset_uri


BASE_DIR = Path(__file__).resolve().parent
GRAPH_DIR = BASE_DIR / "__pycache__"
GRAPH_DIR.mkdir(exist_ok=True)

GLOBAL_RELATION_STATE_KEY = "selected_pair_key"
DEFAULT_PERSON_NETWORK_LIMIT = 12
DEFAULT_OVERVIEW_PAIR_LIMIT = 28


@dataclass(slots=True)
class PairRecordSample:
    relation_type: str
    raw_relation_type: str
    llm_suggested_relation_type: str
    display_status: str
    weight: float
    evidence_ref: str
    context: str
    llm_reason: str = ""
    llm_confidence: float | None = None


@dataclass(slots=True)
class PairProfile:
    pair_key: str
    person_a_id: str
    person_a_name: str
    person_b_id: str
    person_b_name: str
    relation_types: list[str] = field(default_factory=list)
    raw_relation_types: list[str] = field(default_factory=list)
    llm_suggestions: list[str] = field(default_factory=list)
    relation_count: int = 0
    max_weight: float = 0.0
    total_weight: float = 0.0
    display_status_counts: dict[str, int] = field(default_factory=dict)
    evidence_samples: list[str] = field(default_factory=list)
    context_samples: list[str] = field(default_factory=list)
    records: list[PairRecordSample] = field(default_factory=list)

    @property
    def formal_count(self) -> int:
        return int(self.display_status_counts.get("formal", 0))

    @property
    def review_count(self) -> int:
        return int(self.display_status_counts.get("review", 0))


def _unique_texts(values: pd.Series, limit: int | None = None) -> list[str]:
    items: list[str] = []
    for value in values.astype(str):
        text = value.strip()
        if text and text not in items:
            items.append(text)
        if limit and len(items) >= limit:
            break
    return items


def _unique_join(values: pd.Series, limit: int = 4) -> str:
    items = _unique_texts(values, limit=limit)
    return " / ".join(items) if items else "未标注"


def _parse_confidence(value: object) -> float | None:
    text = str(value).strip().lower()
    if not text:
        return None
    if text == "low":
        return 0.40
    if text == "medium":
        return 0.70
    if text == "high":
        return 0.90
    try:
        return float(text)
    except ValueError:
        return None


def _resolved_llm_suggestion(row: pd.Series) -> str:
    suggestion = str(row.get("llm_suggested_relation_type", "")).strip()
    if suggestion:
        return suggestion
    final_relation = str(row.get("final_relation_type", "")).strip()
    if final_relation:
        return final_relation
    relation_family = str(row.get("Relation_Family", "")).strip()
    if relation_family:
        return relation_family
    return str(row.get("Relation_Type", "")).strip()


def _resolved_llm_reason(row: pd.Series) -> str:
    reason = str(row.get("llm_reason", "")).strip()
    if reason:
        return reason
    correction_reason = str(row.get("correction_reason", "")).strip()
    if correction_reason:
        return correction_reason
    raw_relation = str(row.get("raw_relation_type", "")).strip()
    final_relation = str(row.get("final_relation_type", "")).strip() or str(row.get("Relation_Type", "")).strip()
    if raw_relation and final_relation and raw_relation != final_relation:
        return f"结合证据语境，关系类型由“{raw_relation}”调整为“{final_relation}”。"
    if final_relation:
        return f"现有证据可支持“{final_relation}”判断。"
    return ""


def _resolved_llm_confidence(row: pd.Series) -> float | None:
    for field in ("llm_confidence", "LLM_Confidence", "confidence"):
        confidence = _parse_confidence(row.get(field))
        if confidence is not None:
            return confidence
    return None


def _safe_time_range_label(detail: RelationDetail | None) -> str:
    if detail is None:
        return "时间待补"

    range_label = str(getattr(detail, "time_range_label", "") or "").strip()
    if range_label and range_label != "时间待补":
        return range_label

    first_seen = str(getattr(detail, "first_seen", "") or "").strip()
    last_seen = str(getattr(detail, "last_seen", "") or "").strip()
    if not first_seen and not last_seen:
        return "时间待补"
    if not first_seen:
        first_seen = last_seen
    if not last_seen:
        last_seen = first_seen
    if not first_seen and not last_seen:
        return "时间待补"
    if first_seen == last_seen:
        return first_seen
    return f"{first_seen} 至 {last_seen}"


def _status_label(status: str) -> str:
    mapping = {
        "formal": "正式证据",
        "review": "推断辅助",
        "hidden": "隐藏记录",
    }
    return mapping.get(str(status), str(status) or "未标注")


def _format_status_counts(profile: PairProfile) -> str:
    pieces: list[str] = []
    if profile.formal_count:
        pieces.append(f"正式证据 {profile.formal_count}")
    if profile.review_count:
        pieces.append(f"推断辅助 {profile.review_count}")
    for key, value in profile.display_status_counts.items():
        if key not in {"formal", "review"} and value:
            pieces.append(f"{_status_label(key)} {value}")
    return "｜".join(pieces) if pieces else "状态待补"


def _set_selected_pair(pair_key: str, state_key: str) -> None:
    st.session_state[state_key] = pair_key


def _navigate_to_page(target_page: str, page_state_key: str, page_pending_key: str | None = None) -> None:
    st.session_state[page_state_key] = target_page
    if page_pending_key:
        st.session_state[page_pending_key] = target_page
    st.rerun()


def _render_network_html(net: Network, cache_name: str, height: int) -> None:
    path = GRAPH_DIR / cache_name
    net.save_graph(str(path))
    html = path.read_text(encoding="utf-8")
    html = html.replace(
        "</head>",
        f"""
        <style>
            html, body {{
                margin: 0;
                padding: 0;
                background:
                    linear-gradient(180deg, rgba(247,239,226,.96), rgba(239,226,198,.98)),
                    url("{asset_uri("paper_texture.png")}") center/220px repeat;
                font-family: {CHART_FONT};
            }}
            #mynetwork {{
                width: 100% !important;
                height: {height}px !important;
                border: 1px solid {BORDER};
                background:
                    radial-gradient(circle at 50% 45%, rgba(255,250,241,.62) 0, rgba(243,232,210,.78) 72%, rgba(229,212,183,.92) 100%),
                    url("{asset_uri("paper_texture.png")}") center/220px repeat;
                box-shadow:
                    inset 0 0 0 1px rgba(255,248,236,.82),
                    inset 0 0 26px rgba(117,90,60,.06);
            }}
            .vis-tooltip {{
                padding: .6rem .75rem;
                border: 1px solid {BORDER};
                border-radius: 0;
                background: rgba(248,242,231,.98);
                color: {INK};
                font: 14px/1.65 {CHART_FONT};
                box-shadow: none;
                max-width: 280px;
            }}
        </style>
        </head>
        """,
    )
    components.html(html, height=height + 10)


@st.cache_data(show_spinner=False)
def load_relation_detail_bundle(
    nodes_df: pd.DataFrame,
    edges_df: pd.DataFrame,
) -> dict[str, RelationDetail]:
    return build_relation_detail_index(BASE_DIR, nodes_df, edges_df)


def filter_edges_for_display(edges_df: pd.DataFrame, include_review: bool = False) -> pd.DataFrame:
    statuses = {"formal"}
    if include_review:
        statuses.add("review")
    return edges_df[edges_df["Display_Status"].astype(str).isin(statuses)].copy()


@st.cache_data(show_spinner=False)
def build_pair_profile_index(edges_df: pd.DataFrame) -> dict[str, PairProfile]:
    if edges_df.empty:
        return {}

    working = edges_df.copy()
    if "pair_key" not in working.columns:
        working["pair_key"] = working.apply(lambda row: canonical_pair_key(row["Source"], row["Target"]), axis=1)
    working["_ui_llm_suggestion"] = working.apply(_resolved_llm_suggestion, axis=1)
    working["_ui_llm_reason"] = working.apply(_resolved_llm_reason, axis=1)
    working["_ui_llm_confidence"] = working.apply(_resolved_llm_confidence, axis=1)

    profiles: dict[str, PairProfile] = {}
    for pair_key, group in working.groupby("pair_key", sort=False):
        first_row = group.iloc[0]
        ordered_people = sorted(
            (
                (str(first_row["Source_Name"]), str(first_row["Source"])),
                (str(first_row["Target_Name"]), str(first_row["Target"])),
            ),
            key=lambda item: (item[0], item[1]),
        )
        records: list[PairRecordSample] = []
        for _, row in group.sort_values(["Weight", "Display_Status"], ascending=[False, True]).iterrows():
            records.append(
                PairRecordSample(
                    relation_type=str(row.get("Relation_Family", "")).strip(),
                    raw_relation_type=str(row.get("raw_relation_type", "")).strip(),
                    llm_suggested_relation_type=str(row.get("_ui_llm_suggestion", "")).strip(),
                    display_status=str(row.get("Display_Status", "")).strip() or "formal",
                    weight=float(row.get("Weight", 0) or 0),
                    evidence_ref=str(row.get("Evidence_Ref", "")).strip(),
                    context=str(row.get("Context", "")).strip(),
                    llm_reason=str(row.get("_ui_llm_reason", "")).strip(),
                    llm_confidence=row.get("_ui_llm_confidence"),
                )
            )

        profiles[str(pair_key)] = PairProfile(
            pair_key=str(pair_key),
            person_a_id=ordered_people[0][1],
            person_a_name=ordered_people[0][0],
            person_b_id=ordered_people[1][1],
            person_b_name=ordered_people[1][0],
            relation_types=_unique_texts(group["Relation_Family"], limit=6),
            raw_relation_types=_unique_texts(group["raw_relation_type"], limit=6),
            llm_suggestions=_unique_texts(group["_ui_llm_suggestion"], limit=6),
            relation_count=int(len(group)),
            max_weight=float(group["Weight"].max()),
            total_weight=float(group["Weight"].sum()),
            display_status_counts={str(key): int(value) for key, value in group["Display_Status"].astype(str).value_counts().items()},
            evidence_samples=_unique_texts(group["Evidence_Ref"], limit=3),
            context_samples=_unique_texts(group["Context"], limit=2),
            records=records[:16],
        )
    return profiles


@st.cache_data(show_spinner=False)
def build_pair_summary(edges_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for profile in build_pair_profile_index(edges_df).values():
        rows.append(
            {
                "pair_key": profile.pair_key,
                "人物甲": profile.person_a_name,
                "人物甲ID": profile.person_a_id,
                "人物乙": profile.person_b_name,
                "人物乙ID": profile.person_b_id,
                "relation_types": " / ".join(profile.relation_types) or "未标注",
                "relation_count": profile.relation_count,
                "max_weight": profile.max_weight,
                "formal_count": profile.formal_count,
                "review_count": profile.review_count,
                "evidence": clean_text(profile.evidence_samples[0], 90) if profile.evidence_samples else "暂无",
                "context": clean_text(profile.context_samples[0], 120) if profile.context_samples else "暂无",
            }
        )
    if not rows:
        return pd.DataFrame(
            columns=[
                "pair_key",
                "人物甲",
                "人物甲ID",
                "人物乙",
                "人物乙ID",
                "relation_types",
                "relation_count",
                "max_weight",
                "formal_count",
                "review_count",
                "evidence",
                "context",
            ]
        )
    return pd.DataFrame(rows).sort_values(["relation_count", "max_weight"], ascending=[False, False]).reset_index(drop=True)


@st.cache_data(show_spinner=False)
def build_node_strength_summary(edges_df: pd.DataFrame, limit: int = 8) -> pd.DataFrame:
    if edges_df.empty:
        return pd.DataFrame(columns=["人物", "身份", "关系强度", "关联对象数"])

    source_strength = edges_df.groupby(["Source", "Source_Name", "Source_Role"], as_index=False).agg(
        relation_strength=("Weight", "sum"),
        relation_degree=("Target", "nunique"),
    )
    source_strength = source_strength.rename(columns={"Source": "Id", "Source_Name": "人物", "Source_Role": "身份"})

    target_strength = edges_df.groupby(["Target", "Target_Name", "Target_Role"], as_index=False).agg(
        relation_strength=("Weight", "sum"),
        relation_degree=("Source", "nunique"),
    )
    target_strength = target_strength.rename(columns={"Target": "Id", "Target_Name": "人物", "Target_Role": "身份"})

    combined = pd.concat([source_strength, target_strength], ignore_index=True)
    combined = (
        combined.groupby(["Id", "人物", "身份"], as_index=False)
        .agg(关系强度=("relation_strength", "sum"), 关联对象数=("relation_degree", "sum"))
        .sort_values(["关系强度", "关联对象数"], ascending=[False, False])
        .head(limit)
    )
    combined["关系强度"] = combined["关系强度"].round(1)
    return combined[["人物", "身份", "关系强度", "关联对象数"]]


def person_views(person_id: str, edges_df: pd.DataFrame, events_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    direct = edges_df[(edges_df["Source"] == person_id) | (edges_df["Target"] == person_id)].copy()
    if direct.empty:
        return direct, events_df.iloc[0:0].copy()

    direct["关系人ID"] = direct["Target"].where(direct["Source"] == person_id, direct["Source"])
    direct["关系人"] = direct["Target_Name"].where(direct["Source"] == person_id, direct["Source_Name"])
    direct["关系人身份"] = direct["Target_Role"].where(direct["Source"] == person_id, direct["Source_Role"])
    direct["pair_key"] = direct.apply(lambda row: canonical_pair_key(row["Source"], row["Target"]), axis=1)
    direct = direct.sort_values(["Weight", "关系人"], ascending=[False, True])

    person_events = events_df[events_df["Entity_ID"].apply(lambda value: person_id in split_ids(value))].copy()
    person_events = person_events.sort_values(["Datetime", "Timestamp"], na_position="last")
    return direct, person_events


def build_person_relation_summary(person_name: str, direct_edges: pd.DataFrame) -> pd.DataFrame:
    if direct_edges.empty:
        return pd.DataFrame(
            columns=[
                "pair_key",
                "人物甲",
                "人物乙",
                "关系人ID",
                "关系人身份",
                "relation_types",
                "relation_count",
                "formal_count",
                "review_count",
                "max_weight",
                "evidence",
                "context",
            ]
        )

    summary = (
        direct_edges.groupby(["pair_key", "关系人", "关系人ID", "关系人身份"], as_index=False)
        .agg(
            relation_types=("Relation_Family", lambda series: _unique_join(series, limit=4)),
            relation_count=("Relation_Type", "count"),
            formal_count=("Display_Status", lambda series: int((series.astype(str) == "formal").sum())),
            review_count=("Display_Status", lambda series: int((series.astype(str) == "review").sum())),
            max_weight=("Weight", "max"),
            evidence=("Evidence_Ref", lambda series: clean_text(next((value for value in series if str(value).strip()), ""), 90) or "暂无"),
            context=("Context", lambda series: clean_text(next((value for value in series if str(value).strip()), ""), 120) or "暂无"),
        )
        .sort_values(["relation_count", "max_weight", "关系人"], ascending=[False, False, True])
    )
    summary["人物甲"] = person_name
    summary["人物乙"] = summary["关系人"]
    return summary[
        [
            "pair_key",
            "人物甲",
            "人物乙",
            "关系人ID",
            "关系人身份",
            "relation_types",
            "relation_count",
            "formal_count",
            "review_count",
            "max_weight",
            "evidence",
            "context",
        ]
    ]


def relation_color(label: str) -> str:
    relation = str(label)
    if any(keyword in relation for keyword in ("组织", "领导", "合作", "革命", "同志")):
        return PRIMARY
    if any(keyword in relation for keyword in ("通信", "书信", "师友", "交游", "友人", "友谊")):
        return ACCENT
    if any(keyword in relation for keyword in ("亲属", "婚姻", "家庭")):
        return "#6d4d7a"
    return UMBER


def render_network(person_name: str, person_id: str, direct_edges: pd.DataFrame, default_limit: int = DEFAULT_PERSON_NETWORK_LIMIT) -> None:
    if direct_edges.empty:
        st.info("当前人物没有关系记录。")
        return

    grouped = build_person_relation_summary(person_name, direct_edges)
    expand_key = f"person_network_expand_{person_id}"
    show_all = len(grouped) <= default_limit or st.toggle("查看更多关系", value=False, key=expand_key)
    display_df = grouped if show_all else grouped.head(default_limit)

    if len(grouped) > default_limit and not show_all:
        st.caption(f"首屏仅显示最强 {len(display_df)} 位关系对象，避免人物局部网络首次进入即过载。")

    net = Network(height="560px", width="100%", bgcolor=PAPER_LIGHT, font_color=INK, cdn_resources="in_line")
    net.set_options(
        f"""
        {{
            "layout": {{"randomSeed": 26}},
            "interaction": {{
                "hover": true,
                "tooltipDelay": 90,
                "hideEdgesOnDrag": false,
                "navigationButtons": false,
                "keyboard": false
            }},
            "physics": {{"enabled": false}},
            "nodes": {{
                "shape": "dot",
                "borderWidth": 1.5,
                "borderWidthSelected": 2.2,
                "shadow": false,
                "font": {{
                    "face": "{CHART_FONT}",
                    "size": 17,
                    "color": "{INK}",
                    "strokeWidth": 0
                }}
            }},
            "edges": {{
                "shadow": false,
                "selectionWidth": 0,
                "hoverWidth": 0.3,
                "smooth": {{"enabled": true, "type": "cubicBezier", "roundness": 0.18}},
                "font": {{
                    "face": "{CHART_FONT}",
                    "size": 12,
                    "color": "{INK}",
                    "background": "rgba(247,240,228,0.92)",
                    "strokeWidth": 0
                }}
            }}
        }}
        """
    )
    net.add_node(
        person_id,
        label=person_name,
        x=0,
        y=0,
        size=38,
        fixed=True,
        color={"background": PRIMARY, "border": "#5d1619", "highlight": {"background": PRIMARY, "border": "#3f1114"}},
        font={"face": CHART_FONT, "size": 20, "color": "#f8f0e2"},
        title=f"{escape(person_name)}｜当前人物",
    )

    total = max(len(display_df), 1)
    for index, row in enumerate(display_df.itertuples(index=False), start=1):
        angle = (-math.pi / 2) + (2 * math.pi * (index - 1) / total)
        radius = 265 + min(int(row.relation_count), 6) * 10
        x = math.cos(angle) * radius
        y = math.sin(angle) * radius * 0.82
        edge_color = relation_color(str(row.relation_types))
        label_text = clean_text(str(row.relation_types), 10) if index <= 6 else ""
        node_size = min(20 + float(row.max_weight) * 1.4 + int(row.relation_count) * 1.2, 34)
        edge_width = min(1.8 + float(row.max_weight) * 0.55 + int(row.relation_count) * 0.25, 6.0)
        title = (
            f"{show(row.人物乙)}｜{show(row.关系人身份, '身份待补')}<br>"
            f"关系类型：{show(row.relation_types)}<br>"
            f"关系记录：{int(row.relation_count)}<br>"
            f"正式证据：{int(row.formal_count)}｜推断辅助：{int(row.review_count)}<br>"
            f"证据摘录：{escape(show(row.evidence, '暂无'))}"
        )
        net.add_node(
            str(row.关系人ID),
            label=str(row.人物乙),
            x=x,
            y=y,
            fixed=True,
            size=node_size,
            color={
                "background": PAPER,
                "border": edge_color,
                "highlight": {"background": PAPER_LIGHT, "border": edge_color},
            },
            font={"face": CHART_FONT, "size": 16, "color": INK},
            title=title,
        )
        net.add_edge(
            person_id,
            str(row.关系人ID),
            color={"color": edge_color, "highlight": edge_color, "hover": edge_color, "opacity": 0.88},
            width=edge_width,
            label=label_text,
            font={"face": CHART_FONT, "size": 12, "color": INK, "background": "rgba(247,240,228,0.92)"},
            title=title,
        )

    _render_network_html(net, f"person_{person_id}.html", height=560)


def render_relation_overview_network(pair_df: pd.DataFrame, limit_pairs: int = DEFAULT_OVERVIEW_PAIR_LIMIT) -> None:
    if pair_df.empty:
        st.info("当前条件下暂无可展示的关系网络。")
        return

    top_pairs = pair_df.head(limit_pairs).copy()
    graph = nx.Graph()
    for _, row in top_pairs.iterrows():
        graph.add_edge(
            str(row["人物甲ID"]),
            str(row["人物乙ID"]),
            weight=float(row["max_weight"]),
            relation_types=str(row["relation_types"]),
            relation_count=int(row["relation_count"]),
            evidence=str(row["evidence"]),
            source_name=str(row["人物甲"]),
            target_name=str(row["人物乙"]),
        )
        graph.nodes[str(row["人物甲ID"])]["label"] = str(row["人物甲"])
        graph.nodes[str(row["人物乙ID"])]["label"] = str(row["人物乙"])

    positions = nx.spring_layout(graph, seed=26, weight="weight", k=1.4 / max(math.sqrt(max(graph.number_of_nodes(), 1)), 1))
    weighted_degree = dict(graph.degree(weight="weight"))
    relation_degree = dict(graph.degree())
    max_weighted_degree = max(weighted_degree.values()) if weighted_degree else 1.0

    net = Network(height="580px", width="100%", bgcolor=PAPER_LIGHT, font_color=INK, cdn_resources="in_line")
    net.set_options(
        f"""
        {{
            "layout": {{"randomSeed": 26}},
            "interaction": {{
                "hover": true,
                "tooltipDelay": 90,
                "hideEdgesOnDrag": false,
                "navigationButtons": false,
                "keyboard": false
            }},
            "physics": {{"enabled": false}},
            "nodes": {{
                "shape": "dot",
                "borderWidth": 1.4,
                "font": {{
                    "face": "{CHART_FONT}",
                    "size": 16,
                    "color": "{INK}",
                    "strokeWidth": 0
                }}
            }},
            "edges": {{
                "shadow": false,
                "selectionWidth": 0,
                "hoverWidth": 0.3,
                "smooth": {{"enabled": true, "type": "dynamic", "roundness": 0.12}},
                "font": {{
                    "face": "{CHART_FONT}",
                    "size": 11,
                    "color": "{INK}",
                    "background": "rgba(247,240,228,0.92)",
                    "strokeWidth": 0
                }}
            }}
        }}
        """
    )

    for node_id, attrs in graph.nodes(data=True):
        label = str(attrs.get("label", node_id))
        prominence = weighted_degree.get(node_id, 0.0) / max_weighted_degree if max_weighted_degree else 0.0
        size = 18 + prominence * 16
        border = PRIMARY if prominence >= 0.72 else ACCENT if prominence >= 0.45 else UMBER
        x = float(positions[node_id][0]) * 900
        y = float(positions[node_id][1]) * 620
        net.add_node(
            node_id,
            label=label,
            x=x,
            y=y,
            fixed=True,
            size=size,
            color={
                "background": PAPER,
                "border": border,
                "highlight": {"background": PAPER_LIGHT, "border": border},
            },
            title=f"{escape(label)}<br>关联对象数：{relation_degree.get(node_id, 0)}<br>累计强度：{weighted_degree.get(node_id, 0):.1f}",
        )

    for edge_index, (source, target, attrs) in enumerate(graph.edges(data=True), start=1):
        color = relation_color(str(attrs.get("relation_types", "")))
        label = clean_text(str(attrs.get("relation_types", "")), 10) if edge_index <= 8 else ""
        title = (
            f"{escape(str(attrs.get('source_name', source)))} × {escape(str(attrs.get('target_name', target)))}<br>"
            f"关系类型：{escape(str(attrs.get('relation_types', '未标注')))}<br>"
            f"记录数：{int(attrs.get('relation_count', 0))}｜最高权重：{float(attrs.get('weight', 0)):.1f}<br>"
            f"证据样本：{escape(clean_text(attrs.get('evidence', ''), 80) or '暂无')}"
        )
        net.add_edge(
            source,
            target,
            color={"color": color, "highlight": color, "hover": color, "opacity": 0.85},
            width=min(1.5 + float(attrs.get("weight", 0)) * 0.45, 5.8),
            label=label,
            font={"face": CHART_FONT, "size": 11, "color": INK, "background": "rgba(247,240,228,0.92)"},
            title=title,
        )

    st.caption(f"全局网络默认仅展示当前筛选条件下最强的 {min(len(top_pairs), limit_pairs)} 组人物关系，以突出核心结构与关键节点。")
    _render_network_html(net, "relation_overview.html", height=580)


def render_relation_entry_list(
    entries_df: pd.DataFrame,
    pair_profiles: dict[str, PairProfile],
    title: str,
    state_key: str = GLOBAL_RELATION_STATE_KEY,
    limit: int = 8,
    widget_key_prefix: str = "relation_entry",
) -> None:
    if entries_df.empty:
        return

    available = entries_df[entries_df["pair_key"].astype(str).isin(pair_profiles.keys())].copy()
    if available.empty:
        return

    st.markdown(f"#### {title}")
    for _, row in available.head(limit).iterrows():
        profile = pair_profiles.get(str(row["pair_key"]))
        if profile is None:
            continue
        summary_col, action_col = st.columns([0.8, 0.2])
        summary_col.markdown(
            f"""
            <div class="relation-entry">
                <div class="relation-entry-title">{escape(str(row["人物甲"]))} × {escape(str(row["人物乙"]))}</div>
                <div class="relation-entry-meta">
                    关系类型：{escape(str(row.get("relation_types", "未标注")))}<br>
                    记录数：{int(row.get("relation_count", 0))}　最高权重：{float(row.get("max_weight", 0)):.1f}<br>
                    {_format_status_counts(profile)}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if action_col.button("查看证据链", key=f"{widget_key_prefix}_{state_key}_{row['pair_key']}"):
            _set_selected_pair(str(row["pair_key"]), state_key)

    if len(available) > limit:
        st.caption(f"当前先显示前 {limit} 组关系，更多人物对可在上方表格中继续筛选。")


def render_relation_detail_panel(
    relation_details: dict[str, RelationDetail],
    pair_profiles: dict[str, PairProfile],
    selected_pair_key: str | None = None,
    state_key: str = GLOBAL_RELATION_STATE_KEY,
    widget_key_prefix: str = "relation_detail",
) -> None:
    pair_key = selected_pair_key or st.session_state.get(state_key)
    if not pair_key:
        st.info("请先从关系网络、人物页或关系总览中选择一组人物关系。")
        return

    detail = relation_details.get(pair_key)
    profile = pair_profiles.get(pair_key)
    if detail is None and profile is None:
        st.info("当前人物对暂无可展示的关系详情。")
        return

    person_a_name = detail.person_a_name if detail else profile.person_a_name
    person_b_name = detail.person_b_name if detail else profile.person_b_name
    relation_types = detail.relation_types if detail and detail.relation_types else (profile.relation_types if profile else [])
    widget_prefix = f"{widget_key_prefix}_{state_key}_{pair_key}"

    title_col, close_col = st.columns([0.82, 0.18])
    title_col.markdown(
        f"""
        <div class="relation-detail-head">
            <div class="relation-detail-title">{escape(person_a_name)} × {escape(person_b_name)}</div>
            <div class="relation-detail-subtitle">
                这里把人物对的关系归类、原始记录、LLM 辅助判断、证据摘录与上下文串成同一条阅读链，便于说明“这条边为什么成立”。
            </div>
            {"<div class='provenance-note'>当前条目仍兼容示例证据结构，后续可直接替换为正式史料数据。</div>" if detail and detail.provenance == "mock" else ""}
        </div>
        """,
        unsafe_allow_html=True,
    )
    if close_col.button("收起详情", key=f"{widget_prefix}_close"):
        st.session_state[state_key] = None
        st.rerun()

    top_a, top_b, top_c, top_d = st.columns(4)
    top_a.metric("关系归类", " / ".join(relation_types) or "未标注")
    top_b.metric("最高权重", f"{profile.max_weight:.1f}" if profile else f"{detail.strength:.1f}")
    top_c.metric("关系记录", profile.relation_count if profile else detail.evidence_count)
    top_d.metric("证据条数", detail.evidence_count if detail else len(profile.records))

    meta_a, meta_b, meta_c, meta_d = st.columns(4)
    meta_a.metric("原始关系", " / ".join(profile.raw_relation_types[:3]) if profile and profile.raw_relation_types else "未标注")
    meta_b.metric("LLM 建议", " / ".join(profile.llm_suggestions[:3]) if profile and profile.llm_suggestions else "无")
    meta_c.metric("展示状态", _format_status_counts(profile) if profile else "状态待补")
    meta_d.metric("时间范围", _safe_time_range_label(detail))

    if profile and profile.review_count:
        st.caption("该人物对同时存在正式证据与推断辅助记录，二者已在下方记录表中分开标示。")

    record_rows: list[dict[str, object]] = []
    if profile:
        for record in profile.records:
            confidence_text = f"{record.llm_confidence:.2f}" if record.llm_confidence is not None else ""
            record_rows.append(
                {
                    "关系类型": record.relation_type or "未标注",
                    "原始关系": record.raw_relation_type or "未标注",
                    "LLM建议": record.llm_suggested_relation_type or "无",
                    "展示状态": _status_label(record.display_status),
                    "权重": record.weight,
                    "证据摘录": clean_text(record.evidence_ref, 90) or "暂无",
                    "上下文": clean_text(record.context, 120) or "暂无",
                    "LLM说明": clean_text(record.llm_reason, 90) or "",
                    "LLM置信度": confidence_text,
                }
            )

    if record_rows:
        filter_options = list(dict.fromkeys([row["关系类型"] for row in record_rows if row["关系类型"]]))
        status_options = list(dict.fromkeys([row["展示状态"] for row in record_rows if row["展示状态"]]))
        control_a, control_b = st.columns([1.05, 0.95])
        selected_types = control_a.multiselect(
            "按关系类型筛选记录",
            options=filter_options,
            default=filter_options,
            key=f"{widget_prefix}_types",
        )
        selected_statuses = control_b.multiselect(
            "按展示状态筛选记录",
            options=status_options,
            default=status_options,
            key=f"{widget_prefix}_statuses",
        )
        filtered_records = [
            row
            for row in record_rows
            if (not selected_types or row["关系类型"] in selected_types)
            and (not selected_statuses or row["展示状态"] in selected_statuses)
        ]
        st.markdown("#### 关系记录摘要")
        st.dataframe(pd.DataFrame(filtered_records), width="stretch", hide_index=True)
    else:
        st.info("当前人物对暂无结构化关系记录摘要。")

    if detail is None or not detail.evidence_samples:
        return

    evidence_types = [item.relation_type for item in detail.evidence_samples if item.relation_type]
    evidence_filter_options = list(dict.fromkeys(relation_types + evidence_types))
    selected_evidence_types = st.multiselect(
        "按关系类型筛选证据摘录",
        options=evidence_filter_options,
        default=evidence_filter_options,
        key=f"{widget_prefix}_evidence_types",
    )
    sort_mode = st.selectbox(
        "证据时间顺序",
        ["按时间升序", "按时间降序"],
        key=f"{widget_prefix}_sort",
    )
    expand_all = st.toggle("展开全部证据摘录", value=False, key=f"{widget_prefix}_expand")

    filtered_evidences = [
        item
        for item in detail.evidence_samples
        if not selected_evidence_types or item.relation_type in selected_evidence_types
    ]
    filtered_evidences.sort(key=lambda item: (item.sort_date == "", item.sort_date or item.source_date, item.source_title))
    if sort_mode == "按时间降序":
        filtered_evidences = list(reversed(filtered_evidences))

    st.caption("下列摘录按人物对证据链聚合展示；若需区分正式证据与推断辅助，请以上方关系记录表中的展示状态为准。")
    st.markdown("#### 证据链摘录")
    if not filtered_evidences:
        st.info("暂无符合筛选条件的证据样本。")
        return

    for item in filtered_evidences:
        label = f"{item.source_date or '时间待补'}｜{item.source_title}"
        if item.relation_type:
            label = f"{label}｜{item.relation_type}"
        with st.expander(label, expanded=expand_all):
            st.markdown(
                f"""
                <div class="evidence-summary">
                    <div class="evidence-meta">
                        原始文献标题或档案来源：{escape(item.source_title or "来源待补")}<br>
                        日期：{escape(item.source_date or "时间待补")}<br>
                        页码 / 卷次 / 期号 / 出处编号：{escape(item.citation_ref or "待补")}<br>
                        支持说明：{escape(item.support_note or "待补")}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if item.excerpt:
                st.markdown("**原文摘录**")
                st.markdown(
                    f"<div class='excerpt-block'>{escape(item.excerpt).replace(chr(10), '<br>')}</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.info("暂无原文摘录")


def render_home(
    nodes_df: pd.DataFrame,
    edges_df: pd.DataFrame,
    events_df: pd.DataFrame,
    pair_df: pd.DataFrame,
    pair_profiles: dict[str, PairProfile],
    relation_details: dict[str, RelationDetail],
    page_state_key: str,
    page_pending_key: str,
) -> None:
    st.markdown(
        """
        <div class="hero">
            <small>Digital Humanities Archive</small>
            <h1>左联知识库研究成果展示</h1>
            <div class="muted">
                本站以人物关系网络为主舞台，把证据链、笔名消歧、时空地图与人物档案组织为同一条研究叙事。
                建议先进入关系总览，再沿着证据链回看关系如何成立，最后回到地图与统计页理解空间和整体结构。
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    metric_a, metric_b, metric_c, metric_d = st.columns(4)
    metric_a.metric("人物实体", len(nodes_df))
    metric_b.metric("非隐藏关系", len(edges_df))
    metric_c.metric("历史事件", len(events_df))
    metric_d.metric("关系类型", edges_df["Relation_Family"].replace("", "未标注").nunique())

    route_specs = [
        ("01 主舞台", "关系总览", "先看左联核心网络结构、关键节点与最强人物对。", "进入关系网络"),
        ("02 学术支撑", "人物档案", "从人物入口进入关系证据链、别名信息与事件轨迹。", "进入人物档案"),
        ("03 空间扩展", "事件地图", "查看活动地点、事件聚集与空间迁移线索。", "进入事件地图"),
        ("04 结项总结", "统计分析", "把关系结构和事件阶段收束为研究发现。", "进入统计分析"),
    ]
    for column, (kicker, target_page, body, button_label) in zip(st.columns(4), route_specs):
        with column:
            st.markdown(
                f"""
                <div class="route-card">
                    <div class="route-card-kicker">{escape(kicker)}</div>
                    <div class="route-card-title">{escape(target_page)}</div>
                    <div class="route-card-body">{escape(body)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button(button_label, key=f"home_to_{target_page}", width="stretch"):
                _navigate_to_page(target_page, page_state_key, page_pending_key)

    left, right = st.columns([1.15, 0.85])
    with left:
        st.markdown("### 推荐起点：核心关系对")
        st.dataframe(
            pair_df.head(8)[["人物甲", "人物乙", "relation_types", "relation_count", "max_weight", "evidence"]].rename(
                columns={
                    "relation_types": "关系类型",
                    "relation_count": "记录数",
                    "max_weight": "最高权重",
                    "evidence": "证据样本",
                }
            ),
            width="stretch",
            hide_index=True,
        )
        render_relation_entry_list(
            pair_df,
            pair_profiles,
            title="从最强关系对进入证据链",
            state_key=GLOBAL_RELATION_STATE_KEY,
            limit=6,
            widget_key_prefix="home_recommended_pairs",
        )
    with right:
        st.markdown(
            """
            <div class="section-card">
                <div class="section-card-title">浏览主线</div>
                1. 先看“关系总览”，理解网络结构与关键节点。<br>
                2. 再点进具体人物对，核查证据摘录、原始关系与 LLM 辅助判断。<br>
                3. 最后回到地图与统计页，把关系放入空间与阶段结构中解释。
            </div>
            """,
            unsafe_allow_html=True,
        )
        core_nodes = build_node_strength_summary(edges_df, limit=8)
        st.markdown("#### 关键节点速览")
        st.dataframe(core_nodes, width="stretch", hide_index=True)

    render_relation_detail_panel(
        relation_details=relation_details,
        pair_profiles=pair_profiles,
        selected_pair_key=st.session_state.get(GLOBAL_RELATION_STATE_KEY),
        state_key=GLOBAL_RELATION_STATE_KEY,
        widget_key_prefix="home_detail_panel",
    )


def render_people(
    nodes_df: pd.DataFrame,
    edges_df: pd.DataFrame,
    events_df: pd.DataFrame,
    pair_profiles: dict[str, PairProfile],
    relation_details: dict[str, RelationDetail],
    historical_event_frame: pd.DataFrame,
    historical_event_index: dict[str, HistoricalEvent],
    historical_geojson: dict[str, object],
    page_state_key: str,
    page_pending_key: str,
) -> None:
    st.markdown(
        '<div class="page-note">人物档案页是进入关系网络与事件轨迹的入口层。先确认人物身份与别名，再沿着局部关系网络进入证据链，最后回到事件轨迹理解其历史位置。</div>',
        unsafe_allow_html=True,
    )
    query_col, role_col = st.columns([1.1, 0.9])
    with query_col:
        query = st.text_input("检索人物", placeholder="例如：鲁迅、茅盾、核心领导")
    with role_col:
        role = st.selectbox("身份筛选", ["全部身份"] + sorted(nodes_df["Role"].replace("", "未标注").unique().tolist()))

    candidates = nodes_df.copy()
    if query:
        candidates = candidates[match(candidates["search_text"], query)]
    if role != "全部身份":
        candidates = candidates[candidates["Role"].replace("", "未标注") == role]
    candidates = candidates.sort_values("Label")
    if candidates.empty:
        st.warning("没有找到匹配人物。")
        return

    selected_name = st.selectbox("选择人物", candidates["Label"].tolist(), key="people_selected_name")
    person = nodes_df[nodes_df["Label"] == selected_name].iloc[0]
    direct_all, person_events = person_views(str(person["Id"]), edges_df, events_df)
    show_review = st.checkbox("显示推断辅助关系", value=False, key=f"people_show_review_{person['Id']}")
    direct = filter_edges_for_display(direct_all, include_review=show_review)
    direct_summary = build_person_relation_summary(str(person["Label"]), direct)

    head_left, head_right = st.columns([1.0, 1.0])
    with head_left:
        alias_text = show(person["Alias"], "无")
        st.markdown(
            f"""
            <div class="section-card">
                <div class="section-card-title">{escape(show(person["Label"]))}</div>
                生卒：{escape(show(person["Birth_Death"], "未知"))}<br>
                身份：{escape(show(person["Role"], "未知"))}<br>
                别名 / 笔名：{escape(alias_text)}<br>
                资料可靠度：{int(person["Reliability"])}
            </div>
            """,
            unsafe_allow_html=True,
        )
    with head_right:
        metric_a, metric_b, metric_c, metric_d = st.columns(4)
        metric_a.metric("关系对象", len(direct_summary))
        metric_b.metric("关系记录", len(direct))
        metric_c.metric("正式证据", int((direct["Display_Status"].astype(str) == "formal").sum()) if not direct.empty else 0)
        metric_d.metric("事件数量", len(person_events))
        button_a, button_b = st.columns(2)
        if button_a.button("进入关系总览", key=f"people_go_relations_{person['Id']}", width="stretch"):
            _navigate_to_page("关系总览", page_state_key, page_pending_key)
        if button_b.button("进入事件地图", key=f"people_go_events_{person['Id']}", width="stretch"):
            _navigate_to_page("事件地图", page_state_key, page_pending_key)

    tab_network, tab_evidence, tab_timeline = st.tabs(["关系网络", "关系详情", "事件轨迹"])
    with tab_network:
        if direct.empty:
            st.info("当前筛选条件下没有可展示的关系记录。")
        else:
            graph_col, entry_col = st.columns([1.25, 0.75])
            with graph_col:
                render_network(str(person["Label"]), str(person["Id"]), direct)
            with entry_col:
                render_relation_entry_list(
                    direct_summary,
                    pair_profiles,
                    title="图中关系入口",
                    state_key=GLOBAL_RELATION_STATE_KEY,
                    limit=10,
                    widget_key_prefix=f"people_network_{person['Id']}",
                )
                st.caption("建议先从右侧最强关系进入，再展开下方证据链。")
        render_relation_detail_panel(
            relation_details=relation_details,
            pair_profiles=pair_profiles,
            selected_pair_key=st.session_state.get(GLOBAL_RELATION_STATE_KEY),
            state_key=GLOBAL_RELATION_STATE_KEY,
            widget_key_prefix=f"people_network_detail_{person['Id']}",
        )

    with tab_evidence:
        if direct.empty:
            st.info("当前筛选条件下没有可展示的关系记录。")
        else:
            st.dataframe(
                direct[
                    [
                        "关系人",
                        "关系人身份",
                        "Relation_Family",
                        "raw_relation_type",
                        "llm_suggested_relation_type",
                        "Display_Status",
                        "Weight",
                        "Evidence_Preview",
                        "Context_Preview",
                    ]
                ].rename(
                    columns={
                        "关系人": "关系对象",
                        "关系人身份": "身份",
                        "Relation_Family": "关系类型",
                        "raw_relation_type": "原始关系",
                        "llm_suggested_relation_type": "LLM建议",
                        "Display_Status": "展示状态",
                        "Weight": "权重",
                        "Evidence_Preview": "证据预览",
                        "Context_Preview": "上下文预览",
                    }
                ),
                width="stretch",
                hide_index=True,
            )
            render_relation_entry_list(
                direct_summary,
                pair_profiles,
                title="按人物对进入关系详情",
                state_key=GLOBAL_RELATION_STATE_KEY,
                limit=8,
                widget_key_prefix=f"people_evidence_{person['Id']}",
            )
            render_relation_detail_panel(
                relation_details=relation_details,
                pair_profiles=pair_profiles,
                selected_pair_key=st.session_state.get(GLOBAL_RELATION_STATE_KEY),
                state_key=GLOBAL_RELATION_STATE_KEY,
                widget_key_prefix=f"people_evidence_detail_{person['Id']}",
            )

    with tab_timeline:
        left, right = st.columns([0.9, 1.1])
        person_event_ids = set(person_events["Event_ID"].astype(str).tolist()) if not person_events.empty else set()
        person_map_df = historical_event_frame[historical_event_frame["id"].isin(person_event_ids)].copy()
        with left:
            st.dataframe(
                person_events[["Timestamp", "Event", "Hist_Loc", "Current_Loc", "Subjects"]].rename(
                    columns={
                        "Timestamp": "时间",
                        "Event": "事件",
                        "Hist_Loc": "历史地点",
                        "Current_Loc": "今址",
                        "Subjects": "相关人物",
                    }
                ),
                width="stretch",
                hide_index=True,
            )
        with right:
            selected_event_id = render_event_map(
                person_map_df,
                historical_geojson,
                key_prefix=f"people_map_{person['Id']}",
                height=430,
            )
            render_event_detail(historical_event_index.get(selected_event_id), title="事件详情")


def render_relations(
    edges_df: pd.DataFrame,
    pair_profiles: dict[str, PairProfile],
    relation_details: dict[str, RelationDetail],
) -> None:
    st.markdown(
        '<div class="page-note">关系总览页是本站主舞台。这里先回答“左联的核心网络结构是什么”，再沿着人物对进入原始关系、LLM 辅助判断、证据摘录与上下文。</div>',
        unsafe_allow_html=True,
    )
    show_review = st.checkbox("显示推断辅助关系", value=False, key="relations_show_review")
    filtered_edges = filter_edges_for_display(edges_df, include_review=show_review)
    filtered_pair_df = build_pair_summary(filtered_edges)
    if filtered_pair_df.empty:
        st.info("当前条件下暂无可展示的关系人物对。")
        return

    max_weight_cap = max(1, int(filtered_pair_df["max_weight"].max()))
    keyword_col, type_col, weight_col, count_col = st.columns([1.2, 1.0, 0.9, 0.9])
    with keyword_col:
        keyword = st.text_input("检索关系对", placeholder="按人物名或关系类型搜索")
    with type_col:
        family = st.selectbox(
            "关系类型",
            ["全部类型"] + sorted(filtered_edges["Relation_Family"].replace("", "未标注").unique().tolist()),
        )
    with weight_col:
        min_weight = st.slider("最低权重", 0, max_weight_cap, min(3, max_weight_cap))
    with count_col:
        min_records = st.slider("最少记录数", 1, int(filtered_pair_df["relation_count"].max()), 1)

    display = filtered_pair_df[
        (filtered_pair_df["max_weight"] >= min_weight) & (filtered_pair_df["relation_count"] >= min_records)
    ].copy()
    if keyword:
        display = display[
            match(display["人物甲"], keyword) | match(display["人物乙"], keyword) | match(display["relation_types"], keyword)
        ]
    if family != "全部类型":
        display = display[match(display["relation_types"], family)]

    if display.empty:
        st.info("当前筛选条件下没有匹配的人物关系。")
        return

    top_pair = display.iloc[0]
    top_relation = filtered_edges["Relation_Family"].replace("", "未标注").value_counts().idxmax()
    top_nodes = build_node_strength_summary(filtered_edges, limit=8)

    metric_a, metric_b, metric_c, metric_d = st.columns(4)
    metric_a.metric("可见人物对", len(display))
    metric_b.metric("核心关系类型", top_relation)
    metric_c.metric("当前最强人物对", f"{top_pair['人物甲']} × {top_pair['人物乙']}")
    metric_d.metric("关键节点数", len(top_nodes))

    render_relation_overview_network(display, limit_pairs=DEFAULT_OVERVIEW_PAIR_LIMIT)

    left, right = st.columns([1.05, 0.95])
    with left:
        st.dataframe(
            display[["人物甲", "人物乙", "relation_types", "relation_count", "max_weight", "evidence", "context"]].rename(
                columns={
                    "relation_types": "关系类型",
                    "relation_count": "记录数",
                    "max_weight": "最高权重",
                    "evidence": "证据样本",
                    "context": "上下文",
                }
            ),
            width="stretch",
            hide_index=True,
        )
        render_relation_entry_list(
            display,
            pair_profiles,
            title="关系详情入口",
            state_key=GLOBAL_RELATION_STATE_KEY,
            limit=10,
            widget_key_prefix="relations_overview",
        )
    with right:
        st.markdown(
            """
            <div class="section-card">
                <div class="section-card-title">阅读顺序建议</div>
                先看上方全局网络识别核心节点，再从左侧关系对列表进入具体人物对，最后在下方关系详情里核查原始关系、证据摘录与上下文。
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("#### 关键节点")
        st.dataframe(top_nodes, width="stretch", hide_index=True)
        render_relation_detail_panel(
            relation_details=relation_details,
            pair_profiles=pair_profiles,
            selected_pair_key=st.session_state.get(GLOBAL_RELATION_STATE_KEY),
            state_key=GLOBAL_RELATION_STATE_KEY,
            widget_key_prefix="relations_detail_panel",
        )
