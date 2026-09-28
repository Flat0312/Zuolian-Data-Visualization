"""校验 Phase7 关系证据候选包；仅读取候选输出，不写生产数据。"""
from __future__ import annotations

import argparse
import csv
import hashlib
from collections import Counter
from pathlib import Path

REQUIRED_SUPPORT_FIELDS = ("candidate_source_id", "source_path_or_url", "locator", "quote", "receipt_id")
ALLOWED_SUPPORT = {"support", "associated", "conflict", "insufficient"}


def _read(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def validate(out_dir: Path) -> list[str]:
    selection = _read(out_dir / "phase7_relation_selection.csv")
    candidates = _read(out_dir / "phase7_relation_evidence_candidates.csv")
    receipts = _read(out_dir / "phase7_source_receipts.csv")
    errors: list[str] = []
    selected_ids = [row["relation_id"] for row in selection]
    if len(selection) != 12 or len(set(selected_ids)) != 12:
        errors.append("冻结样本必须恰为12条且 relation_id 唯一")
    if {row["relation_id"] for row in candidates} != set(selected_ids):
        errors.append("每条冻结关系必须有候选结论")
    receipt_ids = {row["receipt_id"] for row in receipts}
    source_families: dict[str, set[str]] = {}
    for row in candidates:
        candidate_id = row["candidate_id"]
        support = row["evidence_support"]
        if support not in ALLOWED_SUPPORT:
            errors.append(f"{candidate_id}: 非法 evidence_support={support!r}")
        if row["review_status"] != "pending_human_review":
            errors.append(f"{candidate_id}: 不是待人工复核状态")
        if support == "support":
            missing = [field for field in REQUIRED_SUPPORT_FIELDS if not row[field].strip()]
            if missing:
                errors.append(f"{candidate_id}: support 缺少 {','.join(missing)}")
        if row["quote"].strip():
            actual_hash = hashlib.sha256(row["quote"].strip().encode("utf-8")).hexdigest()
            if row["quote_sha256"] != actual_hash:
                errors.append(f"{candidate_id}: 引文 SHA-256 不匹配")
        elif row["quote_sha256"].strip():
            errors.append(f"{candidate_id}: 空引文不得填写哈希")
        if row["receipt_id"].strip() and row["receipt_id"] not in receipt_ids:
            errors.append(f"{candidate_id}: 回执不存在")
        if row["candidate_source_id"].strip():
            source_families.setdefault(row["relation_id"], set()).add(row["source_family"])
    if any(len(families) > 2 for families in source_families.values()):
        errors.append("单条关系不得超过2个独立来源族")
    receipt_counts = Counter(row["receipt_id"] for row in receipts)
    if any(count != 1 for count in receipt_counts.values()):
        errors.append("回执 ID 必须唯一")
    for row in receipts:
        if not all(row[field].strip() for field in ("receipt_id", "candidate_source_id", "source_path_or_url", "content_hash")):
            errors.append(f"{row['receipt_id'] or '无ID'}: 回执关键字段为空")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="校验 Phase7 候选包")
    parser.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parents[2] / "research" / "drafts" / "reports")
    args = parser.parse_args()
    errors = validate(args.out_dir)
    if errors:
        print("phase7 candidate validation failed")
        print("\n".join(f"- {error}" for error in errors))
        return 1
    print("phase7 candidate validation passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
