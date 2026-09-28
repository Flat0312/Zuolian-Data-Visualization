from __future__ import annotations

import argparse
import json
import sys
from datetime import UTC, datetime
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from kb_schema import REQUIRED_DATA_FILES, validate_data_dir

try:
    from research.analysis.relation_publish_status import PUBLIC_RELATION_STATUSES as CREDIBLE_RELATION_STATUSES
    from research.analysis.relation_publish_status import assign_relation_status_columns
except ModuleNotFoundError:  # 直接脚本运行时
    from relation_publish_status import PUBLIC_RELATION_STATUSES as CREDIBLE_RELATION_STATUSES
    from relation_publish_status import assign_relation_status_columns

DEFAULT_PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_PUBLISH_DIR = PROJECT_ROOT / "data" / "publish"
DEFAULT_REPORT = PROJECT_ROOT / "research" / "drafts" / "reports" / "phase3_publish_gate_report.md"
PUBLIC_MEMBERSHIP_TYPES = {"confirmed_member", "related_person"}
PUBLIC_RELATION_STATUSES = set(CREDIBLE_RELATION_STATUSES)


def _read(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig").fillna("")


def _split_ids(value: object) -> list[str]:
    return [item.strip() for item in str(value).replace("；", ";").split(";") if item.strip()]


def _recompute_relation_statuses(
    relations: pd.DataFrame, rel_evid: pd.DataFrame | None
) -> pd.DataFrame:
    """返修门禁：derived 状态每次按当前行 + 当前证据全量重算，不透传旧值。

    仅 human_adjudication 的 verified/rejected 予以保留；其余一律重算。
    无证据（associated/pending/空 quote/locator）不得产生 supported。
    """
    evid_by_rel: dict[str, list[dict[str, str]]] = {}
    if rel_evid is not None and not rel_evid.empty and "relation_id" in rel_evid.columns:
        for _, erow in rel_evid.iterrows():
            evid_by_rel.setdefault(str(erow.get("relation_id", "")).strip(), []).append(
                {k: ("" if pd.isna(v) else str(v)) for k, v in erow.to_dict().items()}
            )
    relations = relations.copy()
    statuses: list[str] = []
    origins: list[str] = []
    reviewers: list[str] = []
    reviewed_ats: list[str] = []
    review_notes: list[str] = []
    for _, row in relations.iterrows():
        d = {k: ("" if pd.isna(v) else str(v)) for k, v in row.to_dict().items()}
        existing = {
            "reviewer": d.get("reviewer", ""),
            "reviewed_at": d.get("reviewed_at", ""),
            "review_note": d.get("review_note", ""),
        }
        cols = assign_relation_status_columns(
            d, evid_by_rel.get(str(d.get("relation_id", "")).strip(), []), existing=existing
        )
        statuses.append(cols["publish_status"])
        origins.append(cols["publish_status_origin"])
        reviewers.append(cols["reviewer"])
        reviewed_ats.append(cols["reviewed_at"])
        review_notes.append(cols["review_note"])
    relations["publish_status"] = statuses
    relations["publish_status_origin"] = origins
    relations["reviewer"] = reviewers
    relations["reviewed_at"] = reviewed_ats
    relations["review_note"] = review_notes
    return relations


def _sort_for_stable_output(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame
    key = frame.columns[0]
    try:
        return frame.sort_values(by=[key], kind="mergesort").reset_index(drop=True)
    except Exception:
        return frame.reset_index(drop=True)


def build_publish_data(
    processed_dir: Path, publish_dir: Path, report_path: Path, stamp: bool = False
) -> dict[str, object]:
    processed_dir = Path(processed_dir)
    publish_dir = Path(publish_dir)
    publish_dir.mkdir(parents=True, exist_ok=True)

    tables = {filename: _read(processed_dir / filename) for filename in REQUIRED_DATA_FILES}
    memberships = tables["org_memberships.csv"]
    memberships = memberships[memberships["membership_type"].isin(PUBLIC_MEMBERSHIP_TYPES)].copy()
    public_membership_keys = {
        (str(row["person_id"]).strip(), str(row["organization_id"]).strip())
        for _, row in memberships.iterrows()
    }
    public_org_evidence_ids = {
        evidence_id
        for raw_value in memberships["evidence_ids"]
        for evidence_id in _split_ids(raw_value)
    }
    tables["org_memberships.csv"] = memberships
    tables["org_membership_evidences.csv"] = tables["org_membership_evidences.csv"][
        tables["org_membership_evidences.csv"]["evidence_id"].isin(public_org_evidence_ids)
    ].copy()

    relations = tables["person_relations.csv"]
    # 返修门禁：每次按当前证据全量重算 derived 状态，不透传旧 supported。
    # 仅 human_adjudication 的 verified/rejected 保留；associated/pending、无 quote/locator、
    # critical/high、待核验、low、needs_manual_review=yes 均不得进入公开层。
    _rel_evid_for_gate: pd.DataFrame | None = None
    _evid_path = processed_dir / "relation_evidences.csv"
    if _evid_path.exists():
        try:
            _rel_evid_for_gate = _read(_evid_path)
        except Exception:
            _rel_evid_for_gate = None
    relations = _recompute_relation_statuses(relations, _rel_evid_for_gate)
    tables["person_relations.csv"] = relations[
        relations["publish_status"].astype(str).str.strip().isin(PUBLIC_RELATION_STATUSES)
    ].copy()

    facts = tables["fact_evidences.csv"]
    membership_fact_mask = facts["predicate"] == "organization_membership"
    public_membership_fact_mask = facts.apply(
        lambda row: (str(row["subject_id"]).strip(), str(row["object_value"]).strip()) in public_membership_keys,
        axis=1,
    )
    # 第四批A裁决：rejected 状态的事实证据一律不进入发布层。
    non_rejected_mask = facts["review_status"].astype(str).str.strip() != "rejected"
    tables["fact_evidences.csv"] = facts[
        non_rejected_mask & (~membership_fact_mask | public_membership_fact_mask)
    ].copy()

    # 关系证据：rejected 不进入发布层；且仅保留发布层关系的外键闭合子集
    public_relation_ids = set(
        tables["person_relations.csv"]["relation_id"].astype(str).str.strip().tolist()
    ) if "relation_id" in tables["person_relations.csv"].columns else set()
    if (processed_dir / "relation_evidences.csv").exists():
        rel_evid = _read(processed_dir / "relation_evidences.csv")
        rel_evid = rel_evid[rel_evid["review_status"].astype(str).str.strip() != "rejected"].copy()
        if "relation_id" in rel_evid.columns:
            # 公开关系为 0 时保留空表头（外键闭合），不透传非公开关系的证据。
            rel_evid = rel_evid[rel_evid["relation_id"].astype(str).str.strip().isin(public_relation_ids)].copy()
        tables["relation_evidences.csv"] = rel_evid
    # 来源层级：作品/引文原样透传（不做内容过滤，保证映射完整）
    for _optional in ("source_works.csv", "source_passages.csv"):
        _p = processed_dir / _optional
        if _p.exists():
            tables[_optional] = _read(_p)

    # 字节级幂等：所有表按主键排序后写盘，manifest 键排序，CSV 行终止符固定。
    for _name in list(tables.keys()):
        tables[_name] = _sort_for_stable_output(tables[_name])
    manifest_tables: dict[str, dict[str, int]] = {}
    for filename in sorted(set(list(REQUIRED_DATA_FILES) + list(tables.keys()))):
        if filename not in tables:
            continue
        source_count = len(_read(processed_dir / filename)) if (processed_dir / filename).exists() else len(tables[filename])
        output_count = len(tables[filename])
        tables[filename].to_csv(publish_dir / filename, index=False, encoding="utf-8-sig", lineterminator="\n")
        manifest_tables[filename] = {
            "input": source_count,
            "output": output_count,
            "filtered": source_count - output_count,
        }
    # 来源报告同时输出作品数、引文数、独立来源族数（引用条数≠独立来源作品数）
    _src_pub = tables.get("sources.csv", pd.DataFrame())
    _works_pub = tables.get("source_works.csv", pd.DataFrame())
    if not _src_pub.empty and "source_family" in _src_pub.columns:
        _families = int(_src_pub["source_family"].astype(str).str.strip().replace("", pd.NA).dropna().nunique())
    else:
        _families = 0
    source_summary = {
        "citations": int(manifest_tables.get("sources.csv", {}).get("output", 0)),
        "works": int(len(_works_pub)) if _works_pub is not None else 0,
        "families": _families,
    }

    validation = validate_data_dir(publish_dir)
    research_validation = validate_data_dir(processed_dir)
    # 警告分类：过滤关系后预期产生 vs 真正孤立数据（不得只隐藏 warning）。
    from collections import Counter as _Counter

    _pub_warn_counts = _Counter(i.code for i in validation.warnings)
    _res_warn_counts = _Counter(i.code for i in research_validation.warnings)
    manifest: dict[str, object] = {
        "generated_at": datetime.now(UTC).isoformat() if stamp else "unstamped",
        "stamped": bool(stamp),
        "source_dir": str(processed_dir.resolve()),
        "publish_dir": str(publish_dir.resolve()),
        "rules": {
            "public_membership_types": sorted(PUBLIC_MEMBERSHIP_TYPES),
            "public_relation_statuses": sorted(PUBLIC_RELATION_STATUSES),
            "candidate_and_disputed_memberships": "excluded",
            "rejected_fact_evidences": "excluded",
            "rejected_relation_evidences": "excluded",
            "published_relation_evidences": (
                "fk_closed_subset_retains_original_evidence_support;"
                "only_support_not_rejected_with_locator_and_quote_or_context_gates_publication;"
                "associated_rows_are_source_links_not_support_claims"
            ),
            "non_public_relations": "excluded_inferred_pending_rejected",
            "relation_gate": "derived_recomputed_with_evidence_no_passthrough_except_human_verified_rejected",
            "supported_requires": "support_not_rejected_with_locator_and_quote_or_context",
        },
        "tables": manifest_tables,
        "source_summary": source_summary,
        "research_schema_errors": len(research_validation.errors),
        "research_schema_warnings": len(research_validation.warnings),
        "research_warning_breakdown": dict(_res_warn_counts),
        "schema_errors": len(validation.errors),
        "schema_warnings": len(validation.warnings),
        "publish_warning_breakdown": dict(_pub_warn_counts),
    }
    (publish_dir / "publish_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )

    lines = [
        "# Phase 3 发布门禁报告（返修版）",
        "",
        "发布层由研究层自动生成，研究层原始结论未被删除或覆盖。",
        "关系门禁每次按当前证据全量重算 derived 状态，不透传旧 supported；",
        "仅 human_adjudication 的 verified/rejected 保留（须带 reviewer/reviewed_at/review_note）。",
        "supported 要求 support + 未 rejected + locator + (quote|context)；",
        "associated/pending、无 quote/locator、critical/high、待核验、low、needs_manual_review=yes 均不公开。",
        "",
        "| 数据表 | 输入 | 发布 | 过滤 |",
        "| --- | ---: | ---: | ---: |",
    ]
    for filename, counts in manifest_tables.items():
        lines.append(f"| `{filename}` | {counts['input']} | {counts['output']} | {counts['filtered']} |")
    lines.extend(
        [
            "",
            f"- 研究层 Schema 严重错误：{len(research_validation.errors)}；警告：{len(research_validation.warnings)}"
            f"（{dict(_res_warn_counts)}）。",
            f"- 发布层 Schema 严重错误：{len(validation.errors)}；警告：{len(validation.warnings)}"
            f"（{dict(_pub_warn_counts)}）。",
            "- 发布层警告分类：`isolated_person` 增加主要为过滤非公开关系后预期产生（人物失去公开边）；",
            "`orphan_source` 增加主要为非公开关系证据被过滤后、其来源在发布层暂无公开引用（研究层仍保留）。",
            "- 真正孤立数据（研究层即孤立/孤儿）见研究层警告明细，不得只隐藏 warning。",
            "- 公开组织身份仅保留 `confirmed_member` 与 `related_person`。",
            "- `candidate` 与 `disputed` 仅保留在研究层。",
            "- `fact_evidences.csv` 中 `review_status=rejected` 的事实证据不进入发布层。",
            "- 人物关系仅保留 `publish_status` 为 `verified/supported` 的记录；"
            "`pending_review/inferred/rejected`（含 critical/high、待核验、low、needs_manual_review=yes、"
            "associated/pending 证据、无 quote/locator、证据冲突）仅保留在研究层。",
            "- `relation_evidences.csv` 中 `review_status=rejected` 的关系证据不进入发布层；"
            "且仅保留发布层关系的外键闭合子集（公开关系为 0 时为空表头）。",
            "- 发布层 `relation_evidences.csv` 保留各行的原始 `evidence_support`/`review_status`："
            "只有 `support` 且未 rejected、带 locator 与 quote/context 的行才是公开关系的定级依据；"
            "同表的 `associated`/`pending` 行只为外键闭合而保留，语义是「来源关联」，"
            "**不得当作支持该关系的证据引用**。",
            f"- 来源层级：引文 {source_summary['citations']} 条 / 作品 {source_summary['works']} 种 / "
            f"独立来源族 {source_summary['families']} 个（引用条数≠独立来源作品数，同一来源族不重复计数）。",
        ]
    )
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    if validation.errors:
        raise ValueError(validation.summary())
    return manifest


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="从研究数据生成引用闭合的公开发布数据。")
    parser.add_argument("--processed-dir", type=Path, default=DEFAULT_PROCESSED_DIR)
    parser.add_argument("--publish-dir", type=Path, default=DEFAULT_PUBLISH_DIR)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--stamp", action="store_true", help="显式写入动态 generated_at 时间戳（默认不写，保证字节幂等）")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest = build_publish_data(args.processed_dir, args.publish_dir, args.report, stamp=args.stamp)
    print(json.dumps(manifest["tables"], ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
