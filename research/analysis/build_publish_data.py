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

DEFAULT_PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
DEFAULT_PUBLISH_DIR = PROJECT_ROOT / "data" / "publish"
DEFAULT_REPORT = PROJECT_ROOT / "research" / "drafts" / "reports" / "phase3_publish_gate_report.md"
PUBLIC_MEMBERSHIP_TYPES = {"confirmed_member", "related_person"}
PUBLIC_RELATION_STATUSES = {"formal"}


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
    tables["person_relations.csv"] = relations[relations["display_status"].isin(PUBLIC_RELATION_STATUSES)].copy()

    facts = tables["fact_evidences.csv"]
    membership_fact_mask = facts["predicate"] == "organization_membership"
    public_membership_fact_mask = facts.apply(
        lambda row: (str(row["subject_id"]).strip(), str(row["object_value"]).strip()) in public_membership_keys,
        axis=1,
    )
    tables["fact_evidences.csv"] = facts[~membership_fact_mask | public_membership_fact_mask].copy()

    manifest_tables: dict[str, dict[str, int]] = {}
    for filename in REQUIRED_DATA_FILES:
        source_count = len(_read(processed_dir / filename))
        output_count = len(tables[filename])
        tables[filename].to_csv(publish_dir / filename, index=False, encoding="utf-8-sig")
        manifest_tables[filename] = {
            "input": source_count,
            "output": output_count,
            "filtered": source_count - output_count,
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
        },
        "tables": manifest_tables,
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
