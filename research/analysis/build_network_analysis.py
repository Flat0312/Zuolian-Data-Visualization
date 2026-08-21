"""Phase 6: Social network analysis."""
from __future__ import annotations
import pandas as pd
from pathlib import Path
import json

DATA_DIR = Path(r"D:/1大创/左联知识库项目/data/processed")
REPORT_DIR = Path(r"D:/1大创/左联知识库项目/research/drafts/reports")


def build_network():
    import networkx as nx
    persons = pd.read_csv(DATA_DIR / "persons.csv", encoding="utf-8-sig")
    rels = pd.read_csv(DATA_DIR / "person_relations.csv", encoding="utf-8-sig")
    G = nx.Graph()
    for _, p in persons.iterrows():
        G.add_node(p["person_id"], name=str(p.get("standard_name", "")))
    for _, r in rels.iterrows():
        src, tgt = r["source_person_id"], r["target_person_id"]
        if src in G and tgt in G:
            w = 1.0
            if pd.notna(r.get("weight")):
                try: w = float(r["weight"])
                except: w = 1.0
            G.add_edge(src, tgt, weight=w)
    deg_cent = nx.degree_centrality(G)
    bet_cent = nx.betweenness_centrality(G)
    top_deg = sorted(deg_cent.items(), key=lambda x: -x[1])[:20]
    top_bet = sorted(bet_cent.items(), key=lambda x: -x[1])[:20]
    comps = sorted(nx.connected_components(G), key=len, reverse=True)
    try:
        from networkx.algorithms.community import greedy_modularity_communities
        comms = list(greedy_modularity_communities(G))
    except Exception:
        comms = []
    nm = {p["person_id"]: str(p.get("standard_name", "")) for _, p in persons.iterrows()}
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    lines = ["# Phase 6: Network Analysis", ""]
    lines.append("- Nodes: %d" % G.number_of_nodes())
    lines.append("- Edges: %d" % G.number_of_edges())
    lines.append("- Components: %d" % len(comps))
    if comps:
        lines.append("- Largest: %d (%.1f%%)" % (len(comms[0]) if comms else len(comps[0]), len(comps[0])/G.number_of_nodes()*100))
    lines.append("- Communities: %d" % len(comms))
    lines.append("")
    lines.append("## Degree Top 20")
    lines.append("| Rank | Name | ID | Centrality |")
    lines.append("|------|------|----|-----------|")
    for i, (pid, val) in enumerate(top_deg):
        lines.append("| %d | %s | %s | %.4f |" % (i+1, nm.get(pid, pid), pid, val))
    lines.append("")
    lines.append("## Betweenness Top 20")
    lines.append("| Rank | Name | ID | Centrality |")
    lines.append("|------|------|----|-----------|")
    for i, (pid, val) in enumerate(top_bet):
        lines.append("| %d | %s | %s | %.4f |" % (i+1, nm.get(pid, pid), pid, val))
    lines.append("")
    lines.append("## Communities")
    for ci, c in enumerate(comms[:10]):
        members = [nm.get(n, n) for n in sorted(c)[:8]]
        lines.append("- Community %d (%d people): %s" % (ci+1, len(c), ", ".join(members)))
    (REPORT_DIR / "phase6_network_analysis_report.md").write_text("\n".join(lines), encoding="utf-8")
    top5d = [(nm.get(p, ""), round(v, 4)) for p, v in top_deg[:5]]
    top5b = [(nm.get(p, ""), round(v, 4)) for p, v in top_bet[:5]]
    return {"nodes": G.number_of_nodes(), "edges": G.number_of_edges(), "components": len(comps), "communities": len(comms), "top_degree": top5d, "top_between": top5b}

if __name__ == "__main__":
    r = build_network()
    print(json.dumps(r, ensure_ascii=False, indent=2))