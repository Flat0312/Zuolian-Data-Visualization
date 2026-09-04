"""Agent A 数据治理迁移脚本（幂等，可重跑，返修版）。

范围（Agent A + 本次返修）：
- relation_evidences 逐条由 relation 的 source_ids/context/evidence_ref 拆分，全部 pending/associated（仅首建；已存在则不覆盖）
- person_relations.publish_status + publish_status_origin + reviewer/reviewed_at/review_note
  每次按当前行（类型/风险/置信/复核标记）+ 当前 relation_evidences 全量重算，不透传旧值；
  仅 human_adjudication 的 verified/rejected 予以保留（须带审计字段）
- supported 仅当存在 support + 未 rejected + locator + (quote|context) 证据；associated/pending 候选不得冒充支持
- low/同属组织/空间共现/时空共现不得进入公开层
- sources.source_family + 相对路径化；source_works / source_passages 一律走 sync_source_layer 统一入口

只迁移现有证据和状态，不新增史实判断；不改风险值（无反向降险）。
时空（events/places/participants）归 Agent B，本脚本不触碰。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from research.analysis.relation_publish_status import assign_relation_status_columns
from research.analysis.source_layer import portable_source_path, source_family_for, sync_source_layer

PROCESSED = PROJECT_ROOT / "data" / "processed"

STRENGTH_TO_LEVEL = {"一手": "A", "二手": "B", "转引": "C", "参考": "C", "推断": "D"}


def _read(name: str) -> pd.DataFrame:
    return pd.read_csv(PROCESSED / name, encoding="utf-8-sig", dtype=str).fillna("")


def _write(df: pd.DataFrame, name: str) -> None:
    df.to_csv(PROCESSED / name, index=False, encoding="utf-8-sig")


def migrate_relations() -> dict:
    rel = _read("person_relations.csv")
    evid = _read("relation_evidences.csv") if (PROCESSED / "relation_evidences.csv").exists() else pd.DataFrame()
    evid_by_rel: dict[str, list[dict[str, str]]] = {}
    if not evid.empty:
        for _, erow in evid.iterrows():
            evid_by_rel.setdefault(str(erow.get("relation_id", "")), []).append(erow.to_dict())
    statuses, origins, reviewers, reviewed_ats, review_notes = [], [], [], [], []
    for _, row in rel.iterrows():
        d = row.to_dict()
        cols = assign_relation_status_columns(
            d,
            evid_by_rel.get(str(d.get("relation_id", "")), []),
            existing={
                "reviewer": str(d.get("reviewer", "") or ""),
                "reviewed_at": str(d.get("reviewed_at", "") or ""),
                "review_note": str(d.get("review_note", "") or ""),
            },
        )
        statuses.append(cols["publish_status"])
        origins.append(cols["publish_status_origin"])
        reviewers.append(cols["reviewer"])
        reviewed_ats.append(cols["reviewed_at"])
        review_notes.append(cols["review_note"])
    rel["publish_status"] = statuses
    rel["publish_status_origin"] = origins
    rel["reviewer"] = reviewers
    rel["reviewed_at"] = reviewed_ats
    rel["review_note"] = review_notes
    # 不得反向降险：risk 列原样保留
    _write(rel, "person_relations.csv")
    counts = pd.Series(statuses).value_counts().to_dict()
    origin_counts = pd.Series(origins).value_counts().to_dict()
    return {"total": len(rel), "status_counts": counts, "origin_counts": origin_counts}


def migrate_relation_evidences() -> dict:
    # 幂等保护：证据表已存在则不重建——后续任何证据语义升级（support/reviewed/quote）
    # 都不得被迁移脚本覆盖清除。
    if (PROCESSED / "relation_evidences.csv").exists():
        existing = _read("relation_evidences.csv")
        return {"rows": len(existing), "skipped_existing": True}
    rel = _read("person_relations.csv")
    src = _read("sources.csv")
    level_map = {}
    for _, r in src.iterrows():
        sid = str(r.get("source_id", "")).strip()
        strength = str(r.get("evidence_strength", "")).strip()
        level_map[sid] = STRENGTH_TO_LEVEL.get(strength, "C")
    rows: list[dict] = []
    # 确定性排序保证幂等
    rel_sorted = rel.sort_values(["relation_id"]).reset_index(drop=True)
    counter = 0
    for _, row in rel_sorted.iterrows():
        rid = str(row.get("relation_id", "")).strip()
        evidence_ref = str(row.get("evidence_ref", "")).strip()
        context = str(row.get("context", "")).strip()
        raw_ids = str(row.get("source_ids", "")).replace("；", ";").replace("、", ";")
        sids = [s.strip() for s in raw_ids.split(";") if s.strip()]
        # 去重保序
        seen: set[str] = set()
        uniq: list[str] = []
        for s in sids:
            if s not in seen:
                seen.add(s)
                uniq.append(s)
        for sid in sorted(uniq):
            counter += 1
            _ctx = context[:800]
            rows.append(
                {
                    "relation_evidence_id": f"RELE-{counter:05d}",
                    "relation_id": rid,
                    "source_id": sid,
                    "locator": evidence_ref[:500],
                    "quote": "",
                    "context": _ctx,
                    "quote_or_context": _ctx,
                    "evidence_support": "associated",
                    "source_level": level_map.get(sid, "C"),
                    "review_status": "pending",
                    "reviewer_note": "由 person_relations 迁移生成，未经人工审核；仅表示来源关联，不断言支持强度。",
                }
            )
    df = pd.DataFrame(
        rows,
        columns=[
            "relation_evidence_id",
            "relation_id",
            "source_id",
            "locator",
            "quote",
            "context",
            "quote_or_context",
            "evidence_support",
            "source_level",
            "review_status",
            "reviewer_note",
        ],
    )
    _write(df, "relation_evidences.csv")
    return {"rows": len(df)}


def migrate_sources() -> dict:
    src = _read("sources.csv")
    # 相对路径化 + family
    src["source_path"] = [portable_source_path(v) for v in src["source_path"].tolist()]
    src["source_family"] = [
        source_family_for(t, p, u)
        for t, p, u in zip(src["title"].tolist(), src["source_path"].tolist(), src["source_url"].tolist())
    ]
    _write(src, "sources.csv")

    # 层级表维护一律走统一注册/同步入口：稳定 ID、citation_count 按实际 passage 数重算
    stats = sync_source_layer(PROCESSED)
    stats["families"] = int(src["source_family"].nunique())
    return stats


def main() -> int:
    # 顺序：先证据/来源层（被依赖方），后关系状态（依赖证据重算）。幂等可重跑。
    r2 = migrate_relation_evidences()
    print(f"relation_evidences: {r2}")
    r3 = migrate_sources()
    print(f"sources: {r3}")
    r1 = migrate_relations()
    print(f"relations: {r1}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
