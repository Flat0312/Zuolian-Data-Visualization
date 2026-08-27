"""为第三批已转正的 9 条 conflict 证据维护结构化裁决状态与消解注记。

背景：第三批事件级审核决策（phase2_batch3_event_review_decisions.md，
授权待追认）落地后，这 9 条 evidence_support=conflict 的记录所述差异相对
原始事件记载已成历史。conflict 标记本身保留——它是裁决依据的存档；
adjudication_status=resolved_by_event_correction 使机器可区分
“原始冲突”“已裁决冲突”与“直接支持”。

本脚本职责：
1. 为 FE-EVP3-* 且 conflict 的行回填/校正 adjudication_status（结构判定：
   仅依据 evidence_support 与 origin_evidence_id 前缀，绝不解析 reviewer_note）；
2. 归一化历史注记中的重复措辞（如“冲突冲突”“按签核决定消解”）为新口径文案，
   不删除任何冲突说明内容。

幂等：全部目标行均为“状态正确且无待归一化片段”时早退不写盘。
"""

from __future__ import annotations

import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"

ORIGIN_PREFIX = "FE-EVP3-"
ADJUDICATION_COLUMN = "adjudication_status"
ADJUDICATION_RESOLVED = "resolved_by_event_correction"

# 新口径注记正文（含检测标记词），替换旧版重复表述。
RESOLUTION_MARKER = "裁决状态为 resolved_by_event_correction"
LEGACY_NOTE_BODY = (
    "2026-08-21 第三批裁决落地：本条所述冲突冲突已按签核决定消解"
    "（事件名称/日期/地点已按本证据方向修正），conflict 标记保留作裁决依据存档。"
)
RESOLUTION_NOTE = (
    "2026-08-21 第三批裁决落地：本条所述冲突已按审核决策消解，"
    + RESOLUTION_MARKER
    + "（授权待追认）（事件名称/日期/地点已按本证据方向修正），"
    "conflict 标记保留作裁决依据存档。"
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


def _is_target(row: dict[str, str]) -> bool:
    """结构化判定第三批 conflict 证据；不依赖 reviewer_note 内容推断状态。"""
    return (
        row.get("origin_evidence_id", "").startswith(ORIGIN_PREFIX)
        and row.get("evidence_support") == "conflict"
    )


def _normalize_note(note: str) -> str:
    note = note.replace(LEGACY_NOTE_BODY, RESOLUTION_NOTE)
    if RESOLUTION_MARKER not in note:
        if not note.endswith("。"):
            note += "。"
        note += RESOLUTION_NOTE
    return note


def main() -> None:
    fields, rows = read_csv(DATA / "fact_evidences.csv")
    if ADJUDICATION_COLUMN not in fields:
        fields = list(fields) + [ADJUDICATION_COLUMN]

    targets = [r for r in rows if _is_target(r)]
    status_fixes = [
        r
        for r in targets
        if r.get(ADJUDICATION_COLUMN, "") != ADJUDICATION_RESOLVED
    ]
    note_fixes = [
        r
        for r in targets
        if LEGACY_NOTE_BODY in r["reviewer_note"] or RESOLUTION_MARKER not in r["reviewer_note"]
    ]
    # 非目标行的非法值视为脏数据一并清空，保持列语义干净。
    stray_fixes = [
        r
        for r in rows
        if not _is_target(r) and r.get(ADJUDICATION_COLUMN, "") not in ("", None)
    ]

    if not (status_fixes or note_fixes or stray_fixes):
        print(
            f"无新增：{len(targets)} 条第三批 conflict 证据均已具备结构化裁决状态与归一化注记，跳过写入。"
        )
        return

    for row in targets:
        row[ADJUDICATION_COLUMN] = ADJUDICATION_RESOLVED
        row["reviewer_note"] = _normalize_note(row["reviewer_note"])
    for row in stray_fixes:
        row[ADJUDICATION_COLUMN] = ""

    write_csv(DATA / "fact_evidences.csv", fields, rows)
    print(
        f"fact_evidences: 处理 {len(targets)} 条第三批 conflict 证据"
        f"（状态回填 {len(status_fixes)}、注记归一 {len(note_fixes)}、杂值清理 {len(stray_fixes)}）"
    )


if __name__ == "__main__":
    main()
