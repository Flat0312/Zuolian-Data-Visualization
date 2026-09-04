"""可信网络分析脚本（Agent B，只读生产数据，可重复运行）。

输出四类网络对照：
1. 全量探索网络（全部关系，仅研究探索用）；
2. 可信关系网络（排除 待核验类型 / needs_manual_review=yes / critical-high 风险 / confidence=low；
   若存在 Agent A 方案的 publish_status 列，则以 verified/supported 收窄）；
3. 按关系证据等级与独立来源数加权的可信网络；
4. 1928—1930 / 1931—1933 / 1934—1936 三个时间切片（事件共参与网络）。

敏感性分析：
- 移除待审核关系前后（全量 vs 可信）；
- 降低《鲁迅日记》单一来源权重前后；
- 降低低等级来源（C/D）权重前后；
比较中心性 Top10 与社区结构变化。

重要边界：person_relations 无逐条日期，时间切片基于事件共参与构建，
不能把切片结果解释为该时期完整关系网络（见报告证据局限）。
本脚本不修改任何生产数据；所有结论均为候选，需人工复核。
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import networkx as nx
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DATA_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_OUT_DIR = PROJECT_ROOT / "research" / "drafts" / "reports"

SNAPSHOT_DATE = "2026-09-04"

TIME_SLICES = (("1928-1930", 1928, 1930), ("1931-1933", 1931, 1933), ("1934-1936", 1934, 1936))

LEVEL_FACTOR = {"A": 1.0, "B": 0.7, "C": 0.4, "D": 0.2}
LUXUN_DIARY_FAMILY = "luxun_diary"


def _read(data_dir: Path, name: str) -> pd.DataFrame:
    return pd.read_csv(data_dir / name, encoding="utf-8-sig", dtype=str).fillna("")


def is_trusted_relation(row: dict[str, str]) -> bool:
    """可信关系判定（与 build_core_upgrade_candidates.py 保持同一语义）。"""
    if str(row.get("final_relation_type", "")).strip() == "待核验":
        return False
    if str(row.get("needs_manual_review", "")).strip().lower() == "yes":
        return False
    if str(row.get("relation_risk_level", "")).strip().lower() in ("critical", "high"):
        return False
    if str(row.get("confidence", "")).strip().lower() == "low":
        return False
    status = str(row.get("publish_status", "")).strip()
    if status and status not in ("verified", "supported"):
        return False
    return True


def _relation_weight(row: dict[str, str]) -> float:
    try:
        return float(str(row.get("weight", "")).strip() or 1)
    except ValueError:
        return 1.0


class EvidenceIndex:
    """关系证据索引：relation_id -> (best_level, families, n_evidences)。

    优先读 relation_evidences.csv（Agent A 新表）；缺失时退化为
    person_relations.source_ids + sources.source_family；再缺失时按 source_id 计数。
    """

    def __init__(self, relations: pd.DataFrame, relation_evidences: pd.DataFrame | None, sources: pd.DataFrame) -> None:
        self.fallback_source_ids = not relation_evidences is not None and len(relation_evidences or []) == 0
        self.family_by_source = (
            dict(zip(sources["source_id"], sources["source_family"]))
            if "source_family" in sources.columns
            else {}
        )
        self.family_fallback_note = ""
        self.by_relation: dict[str, dict[str, object]] = {}
        if relation_evidences is not None and len(relation_evidences):
            strength = (
                dict(zip(sources["source_id"], sources["evidence_strength"]))
                if "evidence_strength" in sources.columns
                else {}
            )
            strength_to_level = {"一手": "A", "二手": "B", "转引": "C", "参考": "C", "推断": "D"}
            for row in relation_evidences.to_dict("records"):
                rid = str(row.get("relation_id", "")).strip()
                if not rid:
                    continue
                sid = str(row.get("source_id", "")).strip()
                level = str(row.get("source_level", "")).strip()
                if level not in LEVEL_FACTOR:
                    level = strength_to_level.get(strength.get(sid, ""), "C")
                fam = self.family_by_source.get(sid, sid)
                rec = self.by_relation.setdefault(rid, {"levels": set(), "families": set(), "n": 0})
                rec["levels"].add(level)  # type: ignore[union-attr]
                rec["families"].add(fam)  # type: ignore[union-attr]
                rec["n"] = int(rec["n"]) + 1  # type: ignore[assignment]
            if not self.family_by_source:
                self.family_fallback_note = "sources.source_family 缺失，独立来源按 source_id 计数（上界口径）"
        else:
            for row in relations.to_dict("records"):
                rid = str(row.get("relation_id", "")).strip()
                raw = str(row.get("source_ids", "")).replace("；", ";")
                sids = [s.strip() for s in raw.split(";") if s.strip()]
                fams = {self.family_by_source.get(s, s) for s in sids}
                self.by_relation[rid] = {"levels": set(), "families": fams, "n": len(sids)}
            self.family_fallback_note = (
                "relation_evidences.csv 缺失，证据等级不可得（等级因子按 B 档计），"
                "独立来源按 source_ids 映射" + ("source_family" if self.family_by_source else "（source_id 上界口径）")
            )

    def profile(self, rid: str) -> tuple[str, set[str], int]:
        rec = self.by_relation.get(rid)
        if not rec:
            return "", set(), 0
        levels = {lv for lv in rec["levels"] if lv in LEVEL_FACTOR}  # type: ignore[union-attr]
        if not levels:
            levels = {"B"}  # 无等级信息时按二手档计，不抬高
        best = max(levels, key=lambda lv: LEVEL_FACTOR[lv])
        return best, set(rec["families"]), int(rec["n"])  # type: ignore[arg-type]


def evidence_weight(base: float, best_level: str, families: set[str], n_evidences: int) -> float:
    """加权公式（文档化、确定性）：base × 等级因子 × 来源族因子。

    无任何证据的关系乘 0.3 降权（有来源关联但缺独立证据表）。
    """
    if n_evidences == 0:
        return round(base * 0.3, 6)
    level_f = LEVEL_FACTOR.get(best_level, 0.7)
    family_f = min(1.0 + 0.25 * max(len(families) - 1, 0), 2.0)
    return round(base * level_f * family_f, 6)


def build_graphs(data_dir: Path) -> dict[str, object]:
    persons = _read(data_dir, "persons.csv")
    relations = _read(data_dir, "person_relations.csv")
    events = _read(data_dir, "events.csv")
    parts = _read(data_dir, "event_participants.csv")
    sources = _read(data_dir, "sources.csv")
    name_by_id = dict(zip(persons["person_id"], persons["standard_name"]))

    re_path = data_dir / "relation_evidences.csv"
    relation_evidences = _read(data_dir, "relation_evidences.csv") if re_path.exists() else None
    ev_index = EvidenceIndex(relations, relation_evidences, sources)

    full_edges: dict[tuple[str, str], float] = {}
    trusted_edges: dict[tuple[str, str], float] = {}
    weighted_edges: dict[tuple[str, str], float] = {}
    luxun_only_edges: set[tuple[str, str]] = set()
    low_grade_edges: set[tuple[str, str]] = set()
    n_total = n_trusted = 0
    person_ids = set(persons["person_id"])

    for row in relations.to_dict("records"):
        s, t = row.get("source_person_id", ""), row.get("target_person_id", "")
        if s not in person_ids or t not in person_ids or s == t:
            continue
        n_total += 1
        key = (min(s, t), max(s, t))
        base = _relation_weight(row)
        full_edges[key] = full_edges.get(key, 0.0) + base
        if not is_trusted_relation(row):
            continue
        n_trusted += 1
        best_level, families, n_evid = ev_index.profile(str(row.get("relation_id", "")).strip())
        trusted_edges[key] = trusted_edges.get(key, 0.0) + base
        weighted_edges[key] = weighted_edges.get(key, 0.0) + evidence_weight(base, best_level, families, n_evid)
        if families and families <= {LUXUN_DIARY_FAMILY}:
            luxun_only_edges.add(key)
        if best_level in ("C", "D"):
            low_grade_edges.add(key)

    g_full = _mk_graph(person_ids, full_edges)
    g_trusted = _mk_graph(person_ids, trusted_edges)
    g_weighted = _mk_graph(person_ids, weighted_edges)

    # 《鲁迅日记》单一来源降权 / 低等级来源降权（在加权网络上做敏感性）
    g_luxun_dw = g_weighted.copy()
    for key in luxun_only_edges:
        if g_luxun_dw.has_edge(*key):
            g_luxun_dw[key[0]][key[1]]["weight"] = round(g_luxun_dw[key[0]][key[1]]["weight"] * 0.5, 6)
    g_low_dw = g_weighted.copy()
    for key in low_grade_edges:
        if g_low_dw.has_edge(*key):
            g_low_dw[key[0]][key[1]]["weight"] = round(g_low_dw[key[0]][key[1]]["weight"] * 0.5, 6)

    # 时间切片：事件共参与
    participants_by_event: dict[str, list[str]] = {}
    for row in parts.to_dict("records"):
        eid, pid = row.get("event_id", ""), row.get("person_id", "")
        if eid and pid in person_ids:
            participants_by_event.setdefault(eid, []).append(pid)
    slice_graphs: dict[str, nx.Graph] = {}
    slice_event_counts: dict[str, int] = {}
    for label, y0, y1 in TIME_SLICES:
        edge_count: dict[tuple[str, str], float] = {}
        used = 0
        for row in events.to_dict("records"):
            d = str(row.get("event_date", "")).strip()
            if not (len(d) >= 4 and d[:4].isdigit()):
                continue
            year = int(d[:4])
            if not (y0 <= year <= y1):
                continue
            plist = sorted(set(participants_by_event.get(row.get("event_id", ""), [])))
            if len(plist) < 2:
                continue
            used += 1
            for i in range(len(plist)):
                for j in range(i + 1, len(plist)):
                    key = (plist[i], plist[j])
                    edge_count[key] = edge_count.get(key, 0.0) + 1.0
        slice_graphs[label] = _mk_graph(person_ids, edge_count)
        slice_event_counts[label] = used

    return {
        "name_by_id": name_by_id,
        "graphs": {
            "full": g_full,
            "trusted": g_trusted,
            "weighted": g_weighted,
            "luxun_diary_downweighted": g_luxun_dw,
            "low_grade_downweighted": g_low_dw,
        },
        "slices": slice_graphs,
        "slice_event_counts": slice_event_counts,
        "counts": {
            "relations_total": n_total,
            "relations_trusted": n_trusted,
            "luxun_only_trusted_edges": len(luxun_only_edges),
            "low_grade_trusted_edges": len(low_grade_edges),
            "publish_status_column": "publish_status" in relations.columns,
            "relation_evidences_table": relation_evidences is not None and len(relation_evidences) > 0,
            "source_family_column": "source_family" in sources.columns,
        },
        "family_fallback_note": ev_index.family_fallback_note,
    }


def _mk_graph(person_ids: set[str], edges: dict[tuple[str, str], float]) -> nx.Graph:
    g = nx.Graph()
    g.add_nodes_from(sorted(person_ids))
    for (s, t), w in edges.items():
        if w > 0:
            g.add_edge(s, t, weight=round(w, 6))
    return g


def graph_metrics(g: nx.Graph, name_by_id: dict[str, str], top_n: int = 10) -> dict[str, object]:
    n_nodes = g.number_of_nodes()
    n_edges = g.number_of_edges()
    isolated = sum(1 for _, d in g.degree() if d == 0)
    components = sorted((sorted(c) for c in nx.connected_components(g)), key=len, reverse=True)
    deg = {p: v for p, v in g.degree()}
    deg_c = nx.degree_centrality(g) if n_edges else {p: 0.0 for p in g.nodes()}
    bet_c = nx.betweenness_centrality(g, weight="weight") if n_edges else {p: 0.0 for p in g.nodes()}
    wdeg = {p: d for p, d in g.degree(weight="weight")}

    def _top(score: dict[str, float]) -> list[list[object]]:
        ordered = sorted(score.items(), key=lambda kv: (-kv[1], kv[0]))[:top_n]
        return [[pid, name_by_id.get(pid, pid), round(float(v), 6)] for pid, v in ordered]

    communities: list[list[str]] = []
    if n_edges:
        raw = nx.algorithms.community.greedy_modularity_communities(g, weight="weight")
        communities = sorted((sorted(c) for c in raw), key=lambda c: (-len(c), c))
    return {
        "nodes": n_nodes,
        "edges": n_edges,
        "isolated_nodes": isolated,
        "components": len(components),
        "largest_component": len(components[0]) if components else 0,
        "mean_degree": round(sum(deg.values()) / n_nodes, 6) if n_nodes else 0.0,
        "total_weight": round(float(sum(d for _, d in g.degree(weight="weight"))), 6),
        "top_degree": _top(deg_c),
        "top_weighted_degree": _top(wdeg),
        "top_betweenness": _top(bet_c),
        "communities": [{"index": i + 1, "size": len(c), "members": c} for i, c in enumerate(communities[:8])],
        "community_count": len(communities),
    }


def _overlap(a: list[list[object]], b: list[list[object]]) -> dict[str, object]:
    ids_a = [r[0] for r in a]
    ids_b = [r[0] for r in b]
    common = set(ids_a) & set(ids_b)
    return {
        "jaccard_top10": round(len(common) / max(len(set(ids_a) | set(ids_b)), 1), 6),
        "common_count": len(common),
        "rank_changes": [
            [pid, ids_a.index(pid) + 1, ids_b.index(pid) + 1]
            for pid in sorted(common)
            if ids_a.index(pid) != ids_b.index(pid)
        ],
    }


def build(data_dir: Path, out_dir: Path) -> dict[str, object]:
    ctx = build_graphs(data_dir)
    graphs: dict[str, nx.Graph] = ctx["graphs"]  # type: ignore[assignment]
    name_by_id: dict[str, str] = ctx["name_by_id"]  # type: ignore[assignment]
    metrics = {label: graph_metrics(g, name_by_id) for label, g in graphs.items()}
    slices: dict[str, nx.Graph] = ctx["slices"]  # type: ignore[assignment]
    slice_metrics = {label: graph_metrics(g, name_by_id) for label, g in slices.items()}

    sensitivity = {
        "remove_pending_relations": {
            "before": metrics["full"],
            "after": metrics["trusted"],
            "degree_top10_overlap": _overlap(metrics["full"]["top_degree"], metrics["trusted"]["top_degree"]),  # type: ignore[index]
            "betweenness_top10_overlap": _overlap(metrics["full"]["top_betweenness"], metrics["trusted"]["top_betweenness"]),  # type: ignore[index]
        },
        "luxun_diary_single_source_downweight": {
            "before": metrics["weighted"],
            "after": metrics["luxun_diary_downweighted"],
            "weighted_degree_top10_overlap": _overlap(
                metrics["weighted"]["top_weighted_degree"],  # type: ignore[index]
                metrics["luxun_diary_downweighted"]["top_weighted_degree"],  # type: ignore[index]
            ),
        },
        "low_grade_source_downweight": {
            "before": metrics["weighted"],
            "after": metrics["low_grade_downweighted"],
            "weighted_degree_top10_overlap": _overlap(
                metrics["weighted"]["top_weighted_degree"],  # type: ignore[index]
                metrics["low_grade_downweighted"]["top_weighted_degree"],  # type: ignore[index]
            ),
        },
    }

    payload = {
        "snapshot": SNAPSHOT_DATE,
        "data_state": ctx["counts"],
        "family_fallback_note": ctx["family_fallback_note"],
        "formulas": {
            "trusted_rule": "排除 待核验类型 / needs_manual_review=yes / critical-high 风险 / confidence=low；若有 publish_status 列则仅 verified/supported",
            "evidence_weight": "base × 等级因子(A1.0/B0.7/C0.4/D0.2) × min(1+0.25×(族数-1), 2.0)；无证据关系 ×0.3",
            "luxun_diary_downweight": "证据全部属于 luxun_diary 族的边权重 ×0.5",
            "low_grade_downweight": "最强证据等级为 C/D 的边权重 ×0.5",
            "time_slices": "按事件 event_date 年份把共参与参与者连边，边权=同事件共现次数；关系本身无日期，切片不代表该期完整关系网",
        },
        "networks": {
            "full_exploratory": metrics["full"],
            "trusted": metrics["trusted"],
            "evidence_weighted": metrics["weighted"],
            "luxun_diary_downweighted": metrics["luxun_diary_downweighted"],
            "low_grade_downweighted": metrics["low_grade_downweighted"],
        },
        "time_slices": {
            label: {**slice_metrics[label], "dated_events_used": ctx["slice_event_counts"][label]}  # type: ignore[index]
            for label in slices
        },
        "sensitivity": sensitivity,
    }

    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "trustworthy_network_analysis.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8"
    )
    _write_report(out_dir, payload)
    _write_findings_candidates(out_dir, payload, name_by_id)
    return {
        "relations": ctx["counts"]["relations_total"],
        "trusted": ctx["counts"]["relations_trusted"],
        "edges_full": metrics["full"]["edges"],
        "edges_trusted": metrics["trusted"]["edges"],
        "edges_weighted": metrics["weighted"]["edges"],
        "slice_edges": {k: v["edges"] for k, v in slice_metrics.items()},
    }


def _write_report(out_dir: Path, payload: dict[str, object]) -> None:
    nets = payload["networks"]
    p = out_dir / "trustworthy_network_analysis_report.md"
    lines = [
        "# 可信网络分析报告（Agent B）",
        "",
        f"- 数据快照日期：{payload['snapshot']}（只读；未修改生产数据）",
        f"- 数据状态：{payload['data_state']}",
        f"- 口径降级说明：{payload['family_fallback_note'] or '无（三层证据表齐备）'}",
        "",
        "## 一、三套主网络（数据观察，可复核）",
        "",
        "| 网络 | 节点 | 边 | 连通分量 | 最大分量 | 社区数 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for label, key in (
        ("全量探索网络", "full_exploratory"),
        ("可信关系网络", "trusted"),
        ("证据加权网络", "evidence_weighted"),
    ):
        m = nets[key]  # type: ignore[index]
        lines.append(
            f"| {label} | {m['nodes']} | {m['edges']} | {m['components']} | {m['largest_component']} | {m['community_count']} |"  # type: ignore[index]
        )
    lines += [
        "",
        "### 度中心性 Top10 对照（数据观察）",
        "",
        "| 排名 | 全量 | 可信 | 加权（加权度） |",
        "| --- | --- | --- | --- |",
    ]
    for i in range(10):
        row = []
        for key, field in (
            ("full_exploratory", "top_degree"),
            ("trusted", "top_degree"),
            ("evidence_weighted", "top_weighted_degree"),
        ):
            m = nets[key][field][i] if i < len(nets[key][field]) else ["", "—", ""]  # type: ignore[index]
            row.append(f"{m[1]}（{m[2]}）" if m[0] else "—")
        lines.append(f"| {i + 1} | {row[0]} | {row[1]} | {row[2]} |")
    lines += [
        "",
        "## 二、时间切片（1928—1930 / 1931—1933 / 1934—1936，事件共参与网络）",
        "",
        "| 切片 | 有日期事件 | 节点 | 边 | 分量 | 最大分量 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for label, m in payload["time_slices"].items():  # type: ignore[union-attr]
        lines.append(
            f"| {label} | {m['dated_events_used']} | {m['nodes_with_edge'] if 'nodes_with_edge' in m else m['nodes']} | {m['edges']} | {m['components']} | {m['largest_component']} |"  # type: ignore[index]
        )
    sens = payload["sensitivity"]
    lines += [
        "",
        "## 三、敏感性分析（数据观察）",
        "",
        f"1. **移除待审核关系前后**：Top10 度中心性重合 {sens['remove_pending_relations']['degree_top10_overlap']['jaccard_top10']}（共同 {sens['remove_pending_relations']['degree_top10_overlap']['common_count']}/10）；社区数 {sens['remove_pending_relations']['before']['community_count']} → {sens['remove_pending_relations']['after']['community_count']}；最大分量 {sens['remove_pending_relations']['before']['largest_component']} → {sens['remove_pending_relations']['after']['largest_component']} 人。",  # type: ignore[index]
        f"2. **《鲁迅日记》单一来源降权前后**（加权网络）：Top10 加权度重合 {sens['luxun_diary_single_source_downweight']['weighted_degree_top10_overlap']['jaccard_top10']}。",  # type: ignore[index]
        f"3. **低等级来源（C/D）降权前后**（加权网络）：Top10 加权度重合 {sens['low_grade_source_downweight']['weighted_degree_top10_overlap']['jaccard_top10']}。",  # type: ignore[index]
        "",
        "## 四、历史解释（候选，须与数据观察分开阅读）",
        "",
        "- 可信网络相对全量网络的变化幅度，反映『待核验/待复核/高风险/低置信』关系对既有网络结论的支撑程度；若 Top10 高度重合，说明核心人物的中心地位不依赖低质量关系；若重合下降，说明既有结论部分由低质量关系撑起，**现行 phase6 报告的全量结论须降级表述**。",
        "- 《鲁迅日记》单一来源降权后的排名变化，直接检验『日记来源系统性放大鲁迅圈层连接』这一资料偏差假说。",
        "- 时间切片的边密度差异（1931—1933 通常最密）与左联活动强度和史料存留相关，但不能解读为组织活跃度的实测排名。",
        "",
        "## 五、证据局限（须声明）",
        "",
        "- person_relations 无逐条日期：时间切片基于事件共参与，**切片间与主网络不可直接比较规模**，也不代表该时期完整关系网络。",
        "- 加权公式的因子（等级因子、来源族因子）是本研究设定的启发式口径，不是史料学标准；结论对因子取值敏感。",
        "- 中心性/社区结果只反映当前数据集内的结构位置，**不得写成历史重要性排名或确定史实**（AGENTS.md 边界）。",
        "- 若 relation_evidences / source_family 缺失，独立来源按 source_id 计数属上界口径（见文首降级说明），会低估来源相关性。",
        "",
    ]
    p.write_text("\n".join(lines), encoding="utf-8")


def _write_findings_candidates(out_dir: Path, payload: dict[str, object], name_by_id: dict[str, str]) -> None:
    nets = payload["networks"]
    sens = payload["sensitivity"]
    full_top = nets["full_exploratory"]["top_degree"]  # type: ignore[index]
    trusted_top = nets["trusted"]["top_degree"]  # type: ignore[index]
    luxun_rank_full = next((i + 1 for i, r in enumerate(full_top) if r[0] == "ZLH-001"), None)
    luxun_rank_trusted = next((i + 1 for i, r in enumerate(trusted_top) if r[0] == "ZLH-001"), None)
    lines = [
        "# 更新研究发现候选（Agent B，全部待人工复核）",
        "",
        f"- 生成依据：`trustworthy_network_analysis.json`（快照 {payload['snapshot']}）。",
        "- 本文件是 phase6_research_findings.md 的**候选更新稿**，未经人工复核前不得替代现行版本。",
        "- 每条候选均区分 数据观察 / 历史解释 / 证据局限。",
        "",
        "## 候选一：鲁迅的中心地位在可信网络下是否保持",
        "",
        f"- 数据观察：全量网络度中心性 Top1 为 {full_top[0][1]}（{full_top[0][2]}）；可信网络 Top1 为 {trusted_top[0][1]}（{trusted_top[0][2]}）；鲁迅在全量榜第 {luxun_rank_full} 位、可信榜第 {luxun_rank_trusted} 位。移除待审核关系后 Top10 重合度 {sens['remove_pending_relations']['degree_top10_overlap']['jaccard_top10']}。",  # type: ignore[index]
        "- 历史解释（候选）：若鲁迅在可信网络仍居首位，则『结构性中心』结论可在较严格口径下复现；若位次下降，则现行 phase6 发现一的表述须降级为『全量数据集内的中心』。",
        "- 证据局限：中心性依赖样本与来源结构；《鲁迅日记》降权敏感性结果（重合度 "
        f"{sens['luxun_diary_single_source_downweight']['weighted_degree_top10_overlap']['jaccard_top10']}）需一并报告。",  # type: ignore[index]
        "",
        "## 候选二：社区结构对低质量关系的敏感性",
        "",
        f"- 数据观察：社区数 全量 {nets['full_exploratory']['community_count']} → 可信 {nets['trusted']['community_count']}；最大分量 {nets['full_exploratory']['largest_component']} → {nets['trusted']['largest_component']} 人。",  # type: ignore[index]
        "- 历史解释（候选）：『文学创作—组织/地下—戏剧电影』三线结构是否在可信口径下保持，需对比两套社区成员构成后人工判断；本候选只报告计数变化。",
        "- 证据局限：greedy modularity 对权重口径敏感；社区边界不宜作为史实表述。",
        "",
        "## 候选三：时间切片揭示的活动重心迁移",
        "",
        "- 数据观察：三个切片的边数分别为 " + "、".join(f"{k}={v['edges']}" for k, v in payload['time_slices'].items()) + "（事件共参与口径）。",
        "- 历史解释（候选）：1931—1933 切片若显著更密，与左联在龙华事件后转入地下、活动记载集中化的通说相符；但这是史料存留结构的反映，不是活动强度的实测。",
        "- 证据局限：关系无日期；仅有日期且有≥2名已知参与者的事件才进入切片（数量见 JSON dated_events_used）。",
        "",
        "## 候选四：单一来源依赖度",
        "",
        f"- 数据观察：可信网络中证据全部来自《鲁迅日记》一族的边有 {payload['data_state']['luxun_only_trusted_edges']} 条；最强证据为 C/D 级的边有 {payload['data_state']['low_grade_trusted_edges']} 条。",  # type: ignore[index]
        "- 历史解释（候选）：两类边占比高说明『可信』仍大量依赖单一/低阶来源，补证队列（见核心补证候选包）应优先覆盖这些边关联的人物。",
        "- 证据局限：luxun_diary 族划分依赖 sources.source_family；列缺失时无法计算（JSON data_state 有标记）。",
        "",
    ]
    (out_dir / "updated_research_findings_candidates.md").write_text("\n".join(lines), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="可信网络分析（只读）")
    parser.add_argument("--data-dir", type=Path, default=DEFAULT_DATA_DIR)
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    summary = build(args.data_dir, args.out_dir)
    print(f"trustworthy network analysis: {summary}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
