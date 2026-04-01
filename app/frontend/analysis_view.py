from __future__ import annotations

from html import escape

import altair as alt
import pandas as pd
import streamlit as st

from relation_view import filter_edges_for_display
from research_findings import ResearchFinding, build_research_analysis_bundle
from styles import ACCENT, BORDER, CHART_FONT, INK, MUTED, PAPER_LIGHT, PRIMARY, RULE, UMBER


def archive_axis(title: str | None, label_angle: int = 0) -> alt.Axis:
    return alt.Axis(
        title=title,
        titleFont=CHART_FONT,
        titleFontSize=13,
        titleFontWeight=600,
        titlePadding=10,
        titleColor=INK,
        labelFont=CHART_FONT,
        labelFontSize=12,
        labelColor=MUTED,
        labelPadding=8,
        labelAngle=label_angle,
        domainColor=BORDER,
        domainOpacity=0.5,
        domainWidth=0.8,
        tickColor=BORDER,
        tickOpacity=0.2,
        ticks=False,
        grid=False,
    )


def archive_chart(chart: alt.Chart, title: str) -> alt.Chart:
    return (
        chart.properties(
            height=320,
            padding={"left": 8, "right": 14, "top": 12, "bottom": 8},
            title=alt.TitleParams(
                title,
                anchor="start",
                color=PRIMARY,
                font=CHART_FONT,
                fontSize=21,
                fontWeight=600,
                offset=16,
            ),
        )
        .configure(background=PAPER_LIGHT)
        .configure_view(fill=PAPER_LIGHT, stroke=RULE, strokeWidth=1)
        .configure_axis(labelFont=CHART_FONT, titleFont=CHART_FONT)
        .configure_legend(
            labelFont=CHART_FONT,
            titleFont=CHART_FONT,
            labelColor=INK,
            titleColor=INK,
            orient="top-left",
        )
    )


def render_bar(df: pd.DataFrame, x: str, y: str, title: str, horizontal: bool = False, color: str = PRIMARY) -> None:
    if df.empty:
        st.info("当前条件下暂无可展示的数据。")
        return
    if horizontal:
        chart = (
            alt.Chart(df)
            .mark_bar(size=22, color=color, stroke=UMBER, strokeWidth=0.8)
            .encode(
                x=alt.X(f"{x}:Q", axis=archive_axis("数量")),
                y=alt.Y(f"{y}:N", sort="-x", axis=archive_axis(None)),
                color=alt.value(color),
                tooltip=[y, x],
            )
        )
    else:
        chart = (
            alt.Chart(df)
            .mark_bar(size=26, color=color, stroke=UMBER, strokeWidth=0.8)
            .encode(
                x=alt.X(f"{x}:N", sort="-y", axis=archive_axis(None, label_angle=-24)),
                y=alt.Y(f"{y}:Q", axis=archive_axis("数量")),
                color=alt.value(color),
                tooltip=[x, y],
            )
        )
    st.altair_chart(archive_chart(chart, title), width="stretch")


def render_line(df: pd.DataFrame, x: str, y: str, title: str) -> None:
    if df.empty:
        st.info("当前条件下暂无可展示的数据。")
        return
    chart = (
        alt.Chart(df)
        .mark_line(
            color=ACCENT,
            strokeWidth=3,
            point=alt.OverlayMarkDef(size=78, filled=True, fill=PRIMARY, stroke=PAPER_LIGHT, strokeWidth=1.4),
        )
        .encode(
            x=alt.X(f"{x}:Q", axis=archive_axis("年份")),
            y=alt.Y(f"{y}:Q", axis=archive_axis("事件数")),
            tooltip=[x, y],
        )
    )
    st.altair_chart(archive_chart(chart, title), width="stretch")


def render_stage_chart(df: pd.DataFrame, title: str) -> None:
    if df.empty:
        st.info("当前条件下暂无可展示的数据。")
        return
    chart = (
        alt.Chart(df)
        .mark_bar(size=34, color=PRIMARY, stroke=UMBER, strokeWidth=0.8)
        .encode(
            x=alt.X("时期:N", sort=df["时期"].tolist(), axis=archive_axis(None, label_angle=-16)),
            y=alt.Y("事件数:Q", axis=archive_axis("事件数")),
            color=alt.value(PRIMARY),
            tooltip=["时期", "事件数"],
        )
    )
    st.altair_chart(archive_chart(chart, title), width="stretch")


def render_centrality_chart(df: pd.DataFrame, title: str) -> None:
    if df.empty:
        st.info("当前条件下暂无可展示的数据。")
        return
    plot_df = df.copy()
    plot_df["排序值"] = plot_df["连接度"]
    long_df = plot_df.melt(
        id_vars=["人物", "排序值"],
        value_vars=["连接度", "中介中心性"],
        var_name="指标",
        value_name="数值",
    )
    chart = (
        alt.Chart(long_df)
        .mark_bar(size=15, stroke=UMBER, strokeWidth=0.5)
        .encode(
            x=alt.X("数值:Q", axis=archive_axis("标准化指标")),
            y=alt.Y("人物:N", sort=alt.SortField(field="排序值", order="descending"), axis=archive_axis(None)),
            color=alt.Color(
                "指标:N",
                scale=alt.Scale(domain=["连接度", "中介中心性"], range=[PRIMARY, ACCENT]),
                legend=alt.Legend(title="指标"),
            ),
            tooltip=["人物", "指标", alt.Tooltip("数值:Q", format=".3f")],
        )
    )
    st.altair_chart(archive_chart(chart, title), width="stretch")


def render_analysis_note(title: str, body: str, evidence: str) -> None:
    st.markdown(
        f"""
        <div class="analysis-note">
            <div class="analysis-note-title">{escape(title)}</div>
            <div class="analysis-note-body">{escape(body)}</div>
            <div class="analysis-note-source">对应证据来源：{escape(evidence)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_research_finding_cards(findings: list[ResearchFinding]) -> None:
    if not findings:
        st.info("当前样本尚不足以生成研究发现卡片。")
        return

    columns = st.columns(2)
    for index, finding in enumerate(findings):
        with columns[index % 2]:
            st.markdown(
                f"""
                <div class="finding-card">
                    <div class="finding-card-title">{escape(finding.title)}</div>
                    <div class="finding-card-body">{escape(finding.content)}</div>
                    <div class="finding-card-label">对应证据来源</div>
                    <div class="finding-card-body">{escape(finding.evidence)}</div>
                    <div class="finding-card-label">学术意义简述</div>
                    <div class="finding-card-body">{escape(finding.significance)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


@st.cache_data(show_spinner=False)
def load_analysis_bundle(
    nodes_df: pd.DataFrame,
    edges_df: pd.DataFrame,
    events_df: pd.DataFrame,
):
    return build_research_analysis_bundle(nodes_df, edges_df, events_df)


def render_analysis(
    nodes_df: pd.DataFrame,
    edges_df: pd.DataFrame,
    events_df: pd.DataFrame,
) -> None:
    st.markdown(
        '<div class="page-note">统计页用于把关系结构、事件阶段与中心性指标收束为可复核的研究发现，不追求面板堆叠，而强调结论、证据与解释之间的对应关系。</div>',
        unsafe_allow_html=True,
    )
    show_review = st.checkbox("显示待审核关系", value=False, key="analysis_show_review")
    analysis_edges = filter_edges_for_display(edges_df, include_review=show_review)
    research_bundle = load_analysis_bundle(nodes_df, analysis_edges, events_df)
    st.caption(f"分析样本：人物 {len(nodes_df)}｜关系 {len(analysis_edges)}｜事件 {len(events_df)}")

    finding_map = {finding.key: finding for finding in research_bundle.findings}

    row1_chart, row1_note = st.columns([1.2, 0.8])
    with row1_chart:
        render_bar(
            research_bundle.strength_ranking.rename(columns={"关系强度": "关系强度"}),
            "关系强度",
            "人物",
            "核心人物关系强度排名",
            horizontal=True,
            color=PRIMARY,
        )
    with row1_note:
        evidence = finding_map.get("core_node").evidence if "core_node" in finding_map else "图表与样本数据。"
        render_analysis_note("发现说明", research_bundle.chart_notes.get("strength", "当前样本尚未形成核心节点判断。"), evidence)

    row2_chart, row2_note = st.columns([1.2, 0.8])
    with row2_chart:
        render_bar(research_bundle.relation_distribution, "关系类型", "记录数", "不同关系类型分布", color=ACCENT)
    with row2_note:
        evidence = finding_map.get("dominant_relation").evidence if "dominant_relation" in finding_map else "图表与样本数据。"
        render_analysis_note("发现说明", research_bundle.chart_notes.get("relation", "当前样本尚未形成关系类型判断。"), evidence)

    row3_chart, row3_note = st.columns([1.2, 0.8])
    with row3_chart:
        render_stage_chart(research_bundle.stage_counts, "各时期事件数量变化")
    with row3_note:
        evidence = finding_map.get("active_period").evidence if "active_period" in finding_map else "图表与样本数据。"
        render_analysis_note("发现说明", research_bundle.chart_notes.get("period", "当前样本尚未形成时期变化判断。"), evidence)

    row4_chart, row4_note = st.columns([1.2, 0.8])
    with row4_chart:
        render_centrality_chart(research_bundle.centrality_compare, "关键人物网络中心性对比")
    with row4_note:
        evidence = finding_map.get("bridge_actor").evidence if "bridge_actor" in finding_map else "图表与样本数据。"
        render_analysis_note("发现说明", research_bundle.chart_notes.get("centrality", "当前样本尚未形成桥梁人物判断。"), evidence)

    st.markdown("### 研究发现")
    render_research_finding_cards(research_bundle.findings)
