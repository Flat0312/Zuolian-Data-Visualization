from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import networkx as nx
import pandas as pd


STAGE_DEFINITIONS: list[tuple[str, int, int]] = [
    ("萌芽与汇流（1922-1927）", 1922, 1927),
    ("组织化推进（1928-1931）", 1928, 1931),
    ("扩散与转折（1932-1934）", 1932, 1934),
    ("战时前夕（1935-1940）", 1935, 1940),
]


@dataclass(slots=True)
class ResearchFinding:
    key: str
    title: str
    content: str
    evidence: str
    significance: str


@dataclass(slots=True)
class ResearchAnalysisBundle:
    strength_ranking: pd.DataFrame
    relation_distribution: pd.DataFrame
    stage_counts: pd.DataFrame
    centrality_compare: pd.DataFrame
    chart_notes: dict[str, str]
    findings: list[ResearchFinding]


def _clean(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return " ".join(str(value).split())


def _safe_years(events_df: pd.DataFrame) -> pd.Series:
    years = pd.to_numeric(events_df.get("Year"), errors="coerce")
    if years.notna().any():
        return years
    timestamps = events_df.get("Timestamp", pd.Series(dtype="object")).fillna("").astype(str)
    return timestamps.str.extract(r"((?:18|19|20)\d{2})")[0].astype(float)


def _stage_label(year: int | float | None) -> str:
    if pd.isna(year):
        return "未定阶段"
    year_value = int(year)
    for label, start, end in STAGE_DEFINITIONS:
        if start <= year_value <= end:
            return label
    return "未定阶段"


def _relation_distribution(edges_df: pd.DataFrame) -> pd.DataFrame:
    relation_counts = edges_df["Relation_Family"].replace("", "未标注").value_counts().reset_index()
    relation_counts.columns = ["关系类型", "记录数"]
    relation_counts["占比"] = relation_counts["记录数"] / relation_counts["记录数"].sum()
    return relation_counts.head(8)


def _strength_ranking(nodes_df: pd.DataFrame, edges_df: pd.DataFrame) -> pd.DataFrame:
    source_strength = edges_df.groupby("Source").agg(
        strength=("Weight", "sum"),
        relation_count=("Target", "count"),
    )
    target_strength = edges_df.groupby("Target").agg(
        strength=("Weight", "sum"),
        relation_count=("Source", "count"),
    )
    combined = source_strength.add(target_strength, fill_value=0).reset_index().rename(columns={"index": "Id", "Source": "Id"})
    if "Id" not in combined.columns:
        combined = combined.rename(columns={combined.columns[0]: "Id"})
    label_map = nodes_df.set_index("Id")["Label"].to_dict()
    role_map = nodes_df.set_index("Id")["Role"].replace("", "未标注").to_dict()
    combined["人物"] = combined["Id"].map(label_map).fillna(combined["Id"])
    combined["角色"] = combined["Id"].map(role_map).fillna("未标注")
    combined["关系强度"] = combined["strength"].round(2)
    combined["关系记录数"] = combined["relation_count"].astype(int)
    return combined.sort_values(["关系强度", "关系记录数"], ascending=[False, False]).head(10)[
        ["Id", "人物", "角色", "关系强度", "关系记录数"]
    ]


def _build_graph(nodes_df: pd.DataFrame, edges_df: pd.DataFrame) -> tuple[nx.Graph, dict[str, str]]:
    label_map = nodes_df.set_index("Id")["Label"].to_dict()
    role_map = nodes_df.set_index("Id")["Role"].replace("", "未标注").to_dict()
    graph = nx.Graph()
    for _, row in nodes_df.iterrows():
        graph.add_node(str(row["Id"]), label=label_map.get(row["Id"], row["Id"]), role=role_map.get(row["Id"], "未标注"))

    # Aggregate repeated pair records first so centrality is computed on the
    # consolidated social graph rather than raw row duplication.
    grouped = (
        edges_df.groupby(["Source", "Target"], as_index=False)
        .agg(weight=("Weight", "sum"), relation_count=("Relation_Type", "count"))
    )
    for _, row in grouped.iterrows():
        source = str(row["Source"])
        target = str(row["Target"])
        weight = float(row["weight"])
        graph.add_edge(
            source,
            target,
            weight=weight,
            relation_count=int(row["relation_count"]),
            distance=1 / max(weight, 0.1),
        )
    return graph, role_map


def _centrality_compare(nodes_df: pd.DataFrame, edges_df: pd.DataFrame) -> pd.DataFrame:
    graph, role_map = _build_graph(nodes_df, edges_df)
    label_map = nodes_df.set_index("Id")["Label"].to_dict()
    degree_centrality = nx.degree_centrality(graph) if graph.number_of_nodes() else {}
    betweenness = nx.betweenness_centrality(graph, weight="distance", normalized=True) if graph.number_of_edges() else {}
    weighted_degree = dict(graph.degree(weight="weight"))

    rows: list[dict[str, Any]] = []
    for node_id in graph.nodes:
        neighbor_roles = {role_map.get(neighbor, "未标注") for neighbor in graph.neighbors(node_id)}
        rows.append(
            {
                "Id": node_id,
                "人物": label_map.get(node_id, node_id),
                "角色": role_map.get(node_id, "未标注"),
                "连接度": degree_centrality.get(node_id, 0.0),
                "中介中心性": betweenness.get(node_id, 0.0),
                "关系强度": float(weighted_degree.get(node_id, 0.0)),
                "角色跨度": len(neighbor_roles),
            }
        )
    centrality_df = pd.DataFrame(rows)
    return centrality_df.sort_values(["连接度", "中介中心性", "关系强度"], ascending=[False, False, False]).head(8)


def _stage_counts(events_df: pd.DataFrame) -> pd.DataFrame:
    years = _safe_years(events_df)
    stage_frame = pd.DataFrame({"Year": years})
    stage_frame = stage_frame.dropna(subset=["Year"]).copy()
    stage_frame["阶段"] = stage_frame["Year"].apply(_stage_label)
    counts = stage_frame["阶段"].value_counts().rename_axis("阶段").reset_index(name="事件数")
    order_map = {label: order for order, (label, _, _) in enumerate(STAGE_DEFINITIONS, start=1)}
    counts["阶段序"] = counts["阶段"].map(order_map).fillna(999).astype(int)
    counts["时期"] = counts["阶段"]
    return counts.sort_values("阶段序")


def _note_core(strength_df: pd.DataFrame) -> tuple[str, ResearchFinding]:
    top = strength_df.iloc[0]
    second = strength_df.iloc[1] if len(strength_df) > 1 else top
    gap = float(top["关系强度"]) - float(second["关系强度"])
    note = (
        f"{top['人物']}在关系强度与直接关联数量上位居前列，说明其不仅参与关系记录最多，"
        f"也更可能承担材料汇聚与网络组织的核心位置。"
    )
    content = (
        f"{top['人物']}以 {float(top['关系强度']):.0f} 的累计关系强度和 {int(top['关系记录数'])} 条直接关系"
        f"居于样本首位；与第二位相比，强度差值为 {gap:.0f}。这一结果提示其是现有材料中最稳定的核心节点。"
    )
    finding = ResearchFinding(
        key="core_node",
        title="核心节点识别",
        content=content,
        evidence=f"图表：核心人物关系强度排名；数据：{top['人物']} 关系强度 {float(top['关系强度']):.0f}，直接关系 {int(top['关系记录数'])}。",
        significance="可用于界定左联关系网络中的中心人物，并为后续文献抽样、个案研究和网络演化分析提供优先观察对象。",
    )
    return note, finding


def _note_relation(relation_df: pd.DataFrame) -> tuple[str, ResearchFinding]:
    top = relation_df.iloc[0]
    share = float(top["占比"]) * 100
    note = (
        f"{top['关系类型']}在全部关系中占比最高，说明现阶段知识库中被记录最多的并非偶发互动，"
        f"而是某一类稳定关系形态。"
    )
    finding = ResearchFinding(
        key="dominant_relation",
        title="主导关系类型",
        content=f"{top['关系类型']}共出现 {int(top['记录数'])} 次，占已归类关系记录的 {share:.1f}%。这一结果表明网络结构主要围绕该类关系展开。",
        evidence=f"图表：不同关系类型分布；数据：{top['关系类型']} {int(top['记录数'])} 条，占比 {share:.1f}%。",
        significance="有助于判断知识库当前更擅长呈现何种社会连接方式，并提醒研究者关注材料来源对关系类型分布的塑形作用。",
    )
    return note, finding


def _note_stage(stage_df: pd.DataFrame) -> tuple[str, ResearchFinding]:
    peak = stage_df.iloc[stage_df["事件数"].idxmax()]
    ordered = stage_df.sort_values("阶段序").reset_index(drop=True)
    if len(ordered) > 1:
        ordered["增量"] = ordered["事件数"].diff().fillna(0)
        surge = ordered.iloc[ordered["增量"].idxmax()]
    else:
        surge = peak
    note = (
        f"{peak['时期']}的事件数量最高，说明这一阶段留下的活动痕迹最为密集；"
        f"若与前一阶段相比增幅明显，则可视为网络活跃度抬升的重要时间段。"
    )
    finding = ResearchFinding(
        key="active_period",
        title="网络活跃期变化",
        content=f"{peak['时期']}记录到 {int(peak['事件数'])} 条事件，为当前样本中的高点；其中 {surge['时期']}相较前一阶段增量最明显，显示事件活动进入集中化阶段。",
        evidence=f"图表：各时期事件数量变化；数据：峰值阶段 {peak['时期']} {int(peak['事件数'])} 条。",
        significance="可为阶段划分、运动史分期和事件簇研究提供经验依据，也有助于把人物关系变化放回到具体历史时间结构中理解。",
    )
    return note, finding


def _note_bridge(centrality_df: pd.DataFrame) -> tuple[str, ResearchFinding]:
    bridge = centrality_df.sort_values(["中介中心性", "角色跨度", "连接度"], ascending=[False, False, False]).iloc[0]
    note = (
        f"{bridge['人物']}的中介中心性最高，说明其更可能位于不同人物群、身份群或关系簇之间，"
        f"承担连接与转译作用，而不只是简单占有大量直接关系。"
    )
    finding = ResearchFinding(
        key="bridge_actor",
        title="桥梁人物识别",
        content=f"{bridge['人物']}在中介中心性上位居首位，且连接到 {int(bridge['角色跨度'])} 类不同角色，显示其可能承担跨群体联结与信息传递的桥梁功能。",
        evidence=f"图表：关键人物网络中心性对比；数据：{bridge['人物']} 中介中心性 {float(bridge['中介中心性']):.3f}，角色跨度 {int(bridge['角色跨度'])}。",
        significance="桥梁人物的识别有助于理解左联网络如何跨越刊物、组织与私人交往等不同场域发生连接。",
    )
    return note, finding


def build_research_analysis_bundle(
    nodes_df: pd.DataFrame,
    edges_df: pd.DataFrame,
    events_df: pd.DataFrame,
) -> ResearchAnalysisBundle:
    # This bundle is deliberately rule-based: metrics come from the data, while
    # textual findings are produced by fixed scholarly templates rather than an online LLM.
    strength_df = _strength_ranking(nodes_df, edges_df)
    relation_df = _relation_distribution(edges_df)
    stage_df = _stage_counts(events_df)
    centrality_df = _centrality_compare(nodes_df, edges_df)

    chart_notes: dict[str, str] = {}
    findings: list[ResearchFinding] = []

    if not strength_df.empty:
        chart_notes["strength"], finding = _note_core(strength_df)
        findings.append(finding)
    if not relation_df.empty:
        chart_notes["relation"], finding = _note_relation(relation_df)
        findings.append(finding)
    if not stage_df.empty:
        chart_notes["period"], finding = _note_stage(stage_df)
        findings.append(finding)
    if not centrality_df.empty:
        chart_notes["centrality"], finding = _note_bridge(centrality_df)
        findings.append(finding)

    return ResearchAnalysisBundle(
        strength_ranking=strength_df,
        relation_distribution=relation_df,
        stage_counts=stage_df,
        centrality_compare=centrality_df,
        chart_notes=chart_notes,
        findings=findings,
    )
