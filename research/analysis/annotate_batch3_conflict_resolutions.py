"""为第三批已转正的 9 条 conflict 证据维护结构化裁决状态 adjudication_status。

背景：第三批事件级审核决策（phase2_batch3_event_review_decisions.md）落地后，
这 9 条 evidence_support=conflict 的记录所述差异相对原始事件记载已成历史。
conflict 标记本身保留——它是裁决依据的存档；adjudication_status=
resolved_by_event_correction 使机器可区分“原始冲突”“已裁决冲突”与“直接支持”。
授权状态：已追认（2026-08-27 项目所有者确认，确认语「已经追认了」，见决策文档第 7 节）。

职责边界（2026-08-27 验收裁决收窄）：
1. 仅维护 FE-EVP3-* 且 conflict 行的 adjudication_status（结构判定：
   仅依据 evidence_support 与 origin_evidence_id 前缀，绝不解析 reviewer_note）；
2. **不改动任何行的 `reviewer_note` 文案**——历次“归一化”改写已被裁决为越界并
   全部回滚至治理起点基线（提交 6e4d8f9）；该约束由
   tests/test_batch3_governance_closeout.py 的逐字节一致性测试永久锁定。

幂等：全部目标行状态正确时早退不写盘。
"""

from __future__ import annotations

import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"

ORIGIN_PREFIX = "FE-EVP3-"
ADJUDICATION_COLUMN = "adjudication_status"
ADJUDICATION_RESOLVED = "resolved_by_event_correction"


def read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        return list(reader.fieldnames or []), list(reader)


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def _is_target(row: dict[str, str]) -> bool:
    """结构化判定第三批 conflict 证据；不依赖 reviewer_note 内容推断状态。"""
    return (
        row.get("origin_evidence_id", "").startswith(ORIGIN_PREFIX)
        and row.get("evidence_support") == "conflict"
    )


def main() -> None:
    fields, rows = read_csv(DATA / "fact_evidences.csv")
    if ADJUDICATION_COLUMN not in fields:
        fields = list(fields) + [ADJUDICATION_COLUMN]

    targets = [r for r in rows if _is_target(r)]
    status_fixes = [
        r for r in targets if r.get(ADJUDICATION_COLUMN, "") != ADJUDICATION_RESOLVED
    ]
    # 非目标行的非法值视为脏数据一并清空，保持列语义干净。
    stray_fixes = [
        r
        for r in rows
        if not _is_target(r) and r.get(ADJUDICATION_COLUMN, "") not in ("", None)
    ]

    if not (status_fixes or stray_fixes):
        print(f"无新增：{len(targets)} 条第三批 conflict 证据均已具备结构化裁决状态，跳过写入。")
        return

    for row in targets:
        row[ADJUDICATION_COLUMN] = ADJUDICATION_RESOLVED
    for row in stray_fixes:
        row[ADJUDICATION_COLUMN] = ""

    write_csv(DATA / "fact_evidences.csv", fields, rows)
    print(
        f"fact_evidences: 处理 {len(targets)} 条第三批 conflict 证据"
        f"（状态回填 {len(status_fixes)}、杂值清理 {len(stray_fixes)}；reviewer_note 不改写）"
    )


if __name__ == "__main__":
    main()
