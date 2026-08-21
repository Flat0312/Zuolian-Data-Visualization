"""为第三批已转正的 9 条 conflict 证据追加"冲突已消解"状态注记。

背景：2026-08-21 第三批人工签核（phase2_batch3_event_review_decisions.md）后，
事件字段已按证据方向修正，这 9 条 evidence_support=conflict 的记录所描述的
差异相对原始事件记载已成历史。conflict 标记本身保留——它是裁决依据的存档，
但需在 reviewer_note 中明确冲突已消解，避免被误读为未决冲突。

幂等：以注记标记词检查防重；全部已注记时早退不写盘。
"""

from __future__ import annotations

import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"

RESOLUTION_MARKER = "冲突已按签核决定消解"
RESOLUTION_NOTE = (
    "2026-08-21 第三批裁决落地：本条所述冲突"
    + RESOLUTION_MARKER
    + "（事件名称/日期/地点已按本证据方向修正），conflict 标记保留作裁决依据存档。"
)


def read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        return list(reader.fieldnames or []), list(reader)


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    fields, rows = read_csv(DATA / "fact_evidences.csv")
    targets = [
        r
        for r in rows
        if r["origin_evidence_id"].startswith("FE-EVP3-")
        and r["evidence_support"] == "conflict"
    ]
    pending_rows = [r for r in targets if RESOLUTION_MARKER not in r["reviewer_note"]]

    if not pending_rows:
        print(f"无新增：{len(targets)} 条第三批 conflict 证据均已注记消解状态，跳过写入。")
        return

    for row in pending_rows:
        if not row["reviewer_note"].endswith("。"):
            row["reviewer_note"] += "。"
        row["reviewer_note"] += RESOLUTION_NOTE

    write_csv(DATA / "fact_evidences.csv", fields, rows)
    print(f"fact_evidences: 注记 {len(pending_rows)}/{len(targets)} 条 conflict 证据（origin FE-EVP3-*）")


if __name__ == "__main__":
    main()
