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
    from research.analysis.relation_publish_status import derive_relation_publish_status
except ModuleNotFoundError:  # 直接脚本运行时
    from relation_publish_status import PUBLIC_RELATION_STATUSES as CREDIBLE_RELATION_STATUSES
    from relation_publish_status import derive_relation_publish_status

DEFAULT_PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_PUBLISH_DIR = PROJECT_ROOT / "data" / "publish"
DEFAULT_REPORT = PROJECT_ROOT / "research" / "drafts" / "reports" / "phase3_publish_gate_report.md"
PUBLIC_MEMBERSHIP_TYPES = {"confirmed_member", "related_person"}
PUBLIC_RELATION_STATUSES = set(CREDIBLE_RELATION_STATUSES)


def _read(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig").fillna("")


def _split_ids(value: object) -> list[str]:
    return [item.strip() for item in str(value).replace("；", ";").split(";") if item.strip()]


def build_publish_data(processed_dir: Path, publish_dir: Path, report_path: Path) -> dict[str, object]:
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
    # 可信度门禁：仅 verified/supported 进入发布层。
    # 兼容旧夹具（无 publish_status 列）：按现有风险/类型/置信派生，不把
    # critical/high、待核验、low、needs_manual_review=yes 当作正式已证实关系展示。
    if "publish_status" not in relations.columns:
        relations = relations.copy()
        relations["publish_status"] = relations.apply(
            lambda row: derive_relation_publish_status(row.to_dict()), axis=1
        )
    else:
        # 缺值回填派生，保证门禁可复现
        mask_empty = relations["publish_status"].astype(str).str.strip() == ""
        if bool(mask_empty.any()):
            relations = relations.copy()
            relations.loc[mask_empty, "publish_status"] = relations[mask_empty].apply(
                lambda row: derive_relation_publish_status(row.to_dict()), axis=1
            )
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
        if public_relation_ids and "relation_id" in rel_evid.columns:
            rel_evid = rel_evid[rel_evid["relation_id"].astype(str).str.strip().isin(public_relation_ids)].copy()
        tables["relation_evidences.csv"] = rel_evid
    # 来源层级：作品/引文原样透传（不做内容过滤，保证映射完整）
    for _optional in ("source_works.csv", "source_passages.csv"):
        _p = processed_dir / _optional
        if _p.exists():
            tables[_optional] = _read(_p)

    manifest_tables: dict[str, dict[str, int]] = {}
    for filename in list(REQUIRED_DATA_FILES) + [f for f in tables if f not in REQUIRED_DATA_FILES]:
        source_count = len(_read(processed_dir / filename)) if (processed_dir / filename).exists() else len(tables[filename])
        output_count = len(tables[filename])
        tables[filename].to_csv(publish_dir / filename, index=False, encoding="utf-8-sig")
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
    manifest: dict[str, object] = {
        "generated_at": datetime.now(UTC).isoformat(),
        "source_dir": str(processed_dir.resolve()),
        "publish_dir": str(publish_dir.resolve()),
        "rules": {
            "public_membership_types": sorted(PUBLIC_MEMBERSHIP_TYPES),
            "public_relation_statuses": sorted(PUBLIC_RELATION_STATUSES),
            "candidate_and_disputed_memberships": "excluded",
            "rejected_fact_evidences": "excluded",
            "rejected_relation_evidences": "excluded",
            "non_public_relations": "excluded_inferred_pending_rejected",
        },
        "tables": manifest_tables,
        "source_summary": source_summary,
        "schema_errors": len(validation.errors),
        "schema_warnings": len(validation.warnings),
    }
    (publish_dir / "publish_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    lines = [
        "# Phase 3 发布门禁报告",
        "",
        "发布层由研究层自动生成，研究层原始结论未被删除或覆盖。",
        "",
        "| 数据表 | 输入 | 发布 | 过滤 |",
        "| --- | ---: | ---: | ---: |",
    ]
    for filename, counts in manifest_tables.items():
        lines.append(f"| `{filename}` | {counts['input']} | {counts['output']} | {counts['filtered']} |")
    lines.extend(
        [
            "",
            f"- Schema 严重错误：{len(validation.errors)}",
            f"- Schema 警告：{len(validation.warnings)}",
            "- 公开组织身份仅保留 `confirmed_member` 与 `related_person`。",
            "- `candidate` 与 `disputed` 仅保留在研究层。",
            "- `fact_evidences.csv` 中 `review_status=rejected` 的事实证据不进入发布层。",
            "- 人物关系仅保留 `publish_status` 为 `verified/supported` 的记录；"
            "`pending_review/inferred/rejected`（含 critical/high、待核验、low、needs_manual_review=yes）仅保留在研究层。",
            "- `relation_evidences.csv` 中 `review_status=rejected` 的关系证据不进入发布层。",
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
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest = build_publish_data(args.processed_dir, args.publish_dir, args.report)
    print(json.dumps(manifest["tables"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
