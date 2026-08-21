"""Phase 5: stratified sample of person_relations for human audit."""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

DATA_DIR = Path(r"D:/1大创/左联知识库项目/data/processed")
REPORT_DIR = Path(r"D:/1大创/左联知识库项目/research/drafts/reports")

SEED = 42
TARGET_SAMPLE = 400


def sample_relations(data_dir: Path = DATA_DIR, seed: int = SEED,
                     target: int = TARGET_SAMPLE) -> dict:
    """Stratified sample by risk_level x relation_type, output review CSV."""
    rels = pd.read_csv(data_dir / "person_relations.csv",
                        encoding="utf-8-sig")

    if "relation_risk_level" not in rels.columns:
        rels["relation_risk_level"] = "unknown"
    rels["relation_risk_level"] = rels["relation_risk_level"].fillna("unknown")
    rels["standard_relation_type"] = rels["standard_relation_type"].fillna("unknown")

    rels["_stratum"] = (rels["relation_risk_level"].astype(str) + "|" +
                        rels["standard_relation_type"].astype(str))

    n = len(rels)
    sample_n = min(target, n)
    sampled = rels.groupby("_stratum", group_keys=False).apply(
        lambda g: g.sample(n=max(1, round(len(g) / n * sample_n)), random_state=seed),
        include_groups=False,
    )
    if len(sampled) > target:
        sampled = sampled.sample(n=target, random_state=seed)

    sampled = sampled.sort_values("relation_id").reset_index(drop=True)

    review = sampled[["relation_id", "source_person_id", "target_person_id",
                       "standard_relation_type", "relation_risk_level",
                       "confidence", "context"]].copy()
    review["human_verdict"] = ""
    review["human_note"] = ""

    report_dir = REPORT_DIR
    report_dir.mkdir(parents=True, exist_ok=True)
    out_csv = report_dir / "phase5_relation_review_template.csv"
    review.to_csv(out_csv, index=False, encoding="utf-8-sig")

    stats = {
        "total_relations": n,
        "sampled": len(sampled),
        "risk_distribution": rels["relation_risk_level"].value_counts().to_dict(),
        "sample_risk_distribution": sampled["relation_risk_level"].value_counts().to_dict(),
        "type_distribution": rels["standard_relation_type"].value_counts().head(10).to_dict(),
    }

    report_path = report_dir / "phase5_relation_audit_report.md"
    lines = ["# Phase 5: 关系数据抽样审计报告", "",
             f"- 总关系数: {n}", f"- 抽样数: {len(sampled)}",
             f"- 抽样率: {len(sampled)/n*100:.1f}%", "",
             "## 风险等级分布（全量）", ""]
    for k, v in stats["risk_distribution"].items():
        lines.append(f"| {k} | {v} |")
    lines += ["", "## 风险等级分布（抽样）", ""]
    for k, v in stats["sample_risk_distribution"].items():
        lines.append(f"| {k} | {v} |")
    lines += ["", "## 关系类型 Top 10", ""]
    for k, v in stats["type_distribution"].items():
        lines.append(f"| {k} | {v} |")
    lines += ["", f"审查模板: {out_csv}"]
    report_path.write_text("\n".join(lines), encoding="utf-8")

    return stats


if __name__ == "__main__":
    stats = sample_relations()
    print(json.dumps(stats, ensure_ascii=False, indent=2))
