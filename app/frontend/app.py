from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from analysis_view import render_analysis
from data_loader import LoadedData, load_data
from event_view import load_historical_map_bundle, render_events
from relation_view import (
    GLOBAL_RELATION_STATE_KEY,
    build_pair_profile_index,
    load_relation_detail_bundle,
    render_home,
    render_people,
    render_relations,
)
from styles import BASE_DIR, apply_style


PAGE_STATE_KEY = "page_nav"
PAGE_WIDGET_KEY = "page_nav_widget"
PAGE_PENDING_KEY = "page_nav_pending"
CACHE_SCHEMA_VERSION_KEY = "cache_schema_version"
CACHE_SCHEMA_VERSION = 2
PAGES = ["首页", "人物档案", "关系总览", "事件地图", "统计分析"]


def initialize_session_state() -> None:
    current_version = st.session_state.get(CACHE_SCHEMA_VERSION_KEY)
    if current_version != CACHE_SCHEMA_VERSION:
        st.cache_data.clear()
        st.session_state[CACHE_SCHEMA_VERSION_KEY] = CACHE_SCHEMA_VERSION
        st.session_state[GLOBAL_RELATION_STATE_KEY] = None

    if PAGE_STATE_KEY not in st.session_state:
        st.session_state[PAGE_STATE_KEY] = "首页"
    if PAGE_WIDGET_KEY not in st.session_state:
        st.session_state[PAGE_WIDGET_KEY] = st.session_state[PAGE_STATE_KEY]
    if PAGE_PENDING_KEY not in st.session_state:
        st.session_state[PAGE_PENDING_KEY] = None
    if GLOBAL_RELATION_STATE_KEY not in st.session_state:
        st.session_state[GLOBAL_RELATION_STATE_KEY] = None


def sync_page_from_widget() -> None:
    st.session_state[PAGE_STATE_KEY] = st.session_state.get(PAGE_WIDGET_KEY, "首页")


def visible_edges(data: LoadedData) -> pd.DataFrame:
    return data.edges[data.edges["Display_Status"].astype(str) != "hidden"].copy()


def pair_summary_from_profiles(pair_profiles: dict) -> pd.DataFrame:
    rows = []
    for profile in pair_profiles.values():
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
                "evidence": profile.evidence_samples[0] if profile.evidence_samples else "暂无",
                "context": profile.context_samples[0] if profile.context_samples else "暂无",
            }
        )
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values(["relation_count", "max_weight"], ascending=[False, False]).reset_index(drop=True)


def render_sidebar(data: LoadedData, visible_edges_df) -> str:
    with st.sidebar:
        st.markdown("## 左联知识库")
        st.caption("以人物关系网络为核心的数字人文成果展示")
        if st.button("刷新最新数据", width="stretch"):
            st.cache_data.clear()
            st.session_state[GLOBAL_RELATION_STATE_KEY] = None
            st.rerun()
        pending_page = st.session_state.get(PAGE_PENDING_KEY)
        if pending_page in PAGES:
            st.session_state[PAGE_STATE_KEY] = pending_page
            st.session_state[PAGE_WIDGET_KEY] = pending_page
            st.session_state[PAGE_PENDING_KEY] = None
        st.radio(
            "页面",
            PAGES,
            key=PAGE_WIDGET_KEY,
            label_visibility="collapsed",
            on_change=sync_page_from_widget,
        )
        current_page = st.session_state.get(PAGE_STATE_KEY, "首页")
        st.markdown("---")
        st.caption(f"人物 {len(data.nodes)} | 关系 {len(visible_edges_df)} | 事件 {len(data.events)}")
    return current_page


def main() -> None:
    st.set_page_config(page_title="左联作家知识库", page_icon="📚", layout="wide")
    apply_style()
    initialize_session_state()

    loading_hint = st.empty()
    loading_hint.info("正在加载数据与视图资源，请稍候...")
    try:
        data = load_data(BASE_DIR)
    except Exception as exc:
        loading_hint.empty()
        st.error("页面初始化失败，请刷新后重试。若在微信内打开，建议选择“在浏览器打开”。")
        st.exception(exc)
        return
    loading_hint.empty()

    visible_edges_df = visible_edges(data)
    page = render_sidebar(data, visible_edges_df)
    selected_pair_key = st.session_state.get(GLOBAL_RELATION_STATE_KEY)

    pair_profiles = build_pair_profile_index(visible_edges_df)
    relation_details = load_relation_detail_bundle(data.nodes, visible_edges_df) if selected_pair_key else {}

    if page == "首页":
        page_loading_hint = st.empty()
        page_loading_hint.info("正在准备首页关系索引...")
        home_pairs = pair_summary_from_profiles(pair_profiles)
        page_loading_hint.empty()
        render_home(
            nodes_df=data.nodes,
            edges_df=visible_edges_df,
            events_df=data.events,
            pair_df=home_pairs,
            pair_profiles=pair_profiles,
            relation_details=relation_details,
            page_state_key=PAGE_STATE_KEY,
            page_pending_key=PAGE_PENDING_KEY,
        )
        return

    if page == "人物档案":
        historical_event_frame, historical_event_index, historical_geojson = load_historical_map_bundle(
            Path(data.data_dir),
            data.nodes,
            data.events,
        )
        render_people(
            nodes_df=data.nodes,
            edges_df=visible_edges_df,
            events_df=data.events,
            pair_profiles=pair_profiles,
            relation_details=relation_details,
            historical_event_frame=historical_event_frame,
            historical_event_index=historical_event_index,
            historical_geojson=historical_geojson,
            page_state_key=PAGE_STATE_KEY,
            page_pending_key=PAGE_PENDING_KEY,
        )
        return

    if page == "关系总览":
        render_relations(
            edges_df=visible_edges_df,
            pair_profiles=pair_profiles,
            relation_details=relation_details,
        )
        return

    if page == "事件地图":
        historical_event_frame, historical_event_index, historical_geojson = load_historical_map_bundle(
            Path(data.data_dir),
            data.nodes,
            data.events,
        )
        render_events(
            nodes_df=data.nodes,
            historical_event_frame=historical_event_frame,
            historical_event_index=historical_event_index,
            historical_geojson=historical_geojson,
        )
        return

    render_analysis(data.nodes, visible_edges_df, data.events)


if __name__ == "__main__":
    main()
