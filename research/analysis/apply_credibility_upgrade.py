"""Agent A 数据治理迁移脚本（幂等，可重跑）。

范围（Agent A）：关系发布状态 + 关系证据 + 来源两层模型。
- person_relations.publish_status 由现有风险/类型/置信/人工标记派生
- relation_evidences 逐条由 relation 的 source_ids/context/evidence_ref 拆分，全部 pending
- sources.source_family + 相对路径化；source_works / source_passages 由现有 sources 分组生成

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

from research.analysis.relation_publish_status import derive_relation_publish_status
from research.analysis.source_layer import file_sha256, portable_source_path, source_family_for

PROCESSED = PROJECT_ROOT / "data" / "processed"

STRENGTH_TO_LEVEL = {"一手": "A", "二手": "B", "转引": "C", "参考": "C", "推断": "D"}


def _read(name: str) -> pd.DataFrame:
    return pd.read_csv(PROCESSED / name, encoding="utf-8-sig", dtype=str).fillna("")


def _write(df: pd.DataFrame, name: str) -> None:
    df.to_csv(PROCESSED / name, index=False, encoding="utf-8-sig")


def migrate_relations() -> dict:
    rel = _read("person_relations.csv")
    # 派生 publish_status：已人工标记 verified/rejected 的予以保留，其余重算
    statuses = []
    for _, row in rel.iterrows():
        d = row.to_dict()
        existing = str(d.get("publish_status", "")).strip()
        if existing in ("verified", "rejected"):
            statuses.append(existing)
        else:
            statuses.append(derive_relation_publish_status(d))
    rel["publish_status"] = statuses
    # 不得反向降险：risk 列原样保留
    _write(rel, "person_relations.csv")
    counts = pd.Series(statuses).value_counts().to_dict()
    return {"total": len(rel), "status_counts": counts}


def migrate_relation_evidences() -> dict:
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

    # works：按 (title, source_path, source_url) 去重
    grouped = src.drop_duplicates(["title", "source_path", "source_url"]).sort_values(
        ["title", "source_path", "source_url"]
    ).reset_index(drop=True)
    # 文件哈希（仅本地存在文件）
    hash_cache: dict[str, str] = {}
    for p in grouped["source_path"].unique().tolist():
        if not p:
            continue
        cand = PROJECT_ROOT / p
        if cand.is_file():
            try:
                hash_cache[p] = file_sha256(cand)
            except Exception:
                hash_cache[p] = ""
    work_rows = []
    work_key_to_id: dict[tuple, str] = {}
    for i, (_, r) in enumerate(grouped.iterrows(), start=1):
        wid = f"WORK-{i:04d}"
        key = (str(r["title"]), str(r["source_path"]), str(r["source_url"]))
        work_key_to_id[key] = wid
        title = str(r["title"])
        author = "鲁迅" if "鲁迅日记" in title else ""
        work_rows.append(
            {
                "work_id": wid,
                "title": title,
                "author": author,
                "version": "",
                "publication_info": "待人工核录",
                "source_category": str(r.get("evidence_type", "")),
                "source_family": str(r.get("source_family", "")),
                "citation_count": int((src["title"] == r["title"]).sum()),
            }
        )
    works_df = pd.DataFrame(
        work_rows,
        columns=["work_id", "title", "author", "version", "publication_info", "source_category", "source_family", "citation_count"],
    )
    works_df.to_csv(PROCESSED / "source_works.csv", index=False, encoding="utf-8-sig")

    # passages：每条 source 一条 passage
    pass_rows = []
    for i, (_, r) in enumerate(src.sort_values("source_id").iterrows(), start=1):
        key = (str(r["title"]), str(r["source_path"]), str(r["source_url"]))
        wid = work_key_to_id.get(key, "")
        spath = str(r.get("source_path", ""))
        fhash = hash_cache.get(spath, "")
        # 仅当文件真实存在才有哈希；网页/空路径留空，不伪造
        pass_rows.append(
            {
                "passage_id": f"PSGN-{i:05d}",
                "work_id": wid,
                "source_id": str(r.get("source_id", "")),
                "locator": str(r.get("citation", ""))[:500],
                "citation": str(r.get("citation", ""))[:800],
                "file_hash": fhash,
                "source_url": str(r.get("source_url", "")),
            }
        )
    pass_df = pd.DataFrame(
        pass_rows,
        columns=["passage_id", "work_id", "source_id", "locator", "citation", "file_hash", "source_url"],
    )
    pass_df.to_csv(PROCESSED / "source_passages.csv", index=False, encoding="utf-8-sig")
    return {
        "sources": len(src),
        "works": len(works_df),
        "passages": len(pass_df),
        "families": int(src["source_family"].nunique()),
    }


def main() -> int:
    r1 = migrate_relations()
    print(f"relations: {r1}")
    r2 = migrate_relation_evidences()
    print(f"relation_evidences: {r2}")
    r3 = migrate_sources()
    print(f"sources: {r3}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
