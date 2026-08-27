"""为第三批已转正的 9 条 conflict 证据维护结构化裁决状态与消解注记。

背景：第三批事件级审核决策（phase2_batch3_event_review_decisions.md）落地后，
这 9 条 evidence_support=conflict 的记录所述差异相对原始事件记载已成历史。
conflict 标记本身保留——它是裁决依据的存档；adjudication_status=
resolved_by_event_correction 使机器可区分“原始冲突”“已裁决冲突”与“直接支持”。
授权状态：**已追认**（2026-08-27 项目所有者确认，见决策文档第 7 节第 6 条）。

本脚本职责：
1. 为 FE-EVP3-* 且 conflict 的行回填/校正 adjudication_status（结构判定：
   仅依据 evidence_support 与 origin_evidence_id 前缀，绝不解析 reviewer_note）；
2. 归一化历史注记措辞至当前授权口径：旧版重复表述（“冲突冲突”“按签核决定
   消解”）与追认前的“（授权待追认）”标记，统一替换为追认后文案；不删除任何
   冲突说明内容。

幂等：全部目标行均为“状态正确且注记已是最新口径”时早退不写盘。
"""

from __future__ import annotations

import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"

ORIGIN_PREFIX = "FE-EVP3-"
ADJUDICATION_COLUMN = "adjudication_status"
ADJUDICATION_RESOLVED = "resolved_by_event_correction"

# 新口径注记正文（含检测标记词），替换历代旧版表述。
RESOLUTION_MARKER = "裁决状态为 resolved_by_event_correction"
AUTH_CONFIRMED_TAG = "（人工授权已于2026-08-27追认，见 phase2_batch3_event_review_decisions.md 第7节）"
RESOLUTION_NOTE = (
    "2026-08-21 第三批裁决落地：本条所述冲突已按审核决策消解，"
    + RESOLUTION_MARKER
    + AUTH_CONFIRMED_TAG
    + "（事件名称/日期/地点已按本证据方向修正），conflict 标记保留作裁决依据存档。"
)

# 历代注记正文 → 当前口径（顺序应用）。v1=收口前重复文案；v2=追认前占位标记。
LEGACY_NOTE_BODIES: list[tuple[str, str]] = [
    (
        "2026-08-21 第三批裁决落地：本条所述冲突冲突已按签核决定消解"
        "（事件名称/日期/地点已按本证据方向修正），conflict 标记保留作裁决依据存档。",
        RESOLUTION_NOTE,
    ),
    ("（授权待追认）", AUTH_CONFIRMED_TAG),
]


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
    for legacy, current in LEGACY_NOTE_BODIES:
        note = note.replace(legacy, current)
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
        r for r in targets if r.get(ADJUDICATION_COLUMN, "") != ADJUDICATION_RESOLVED
    ]
    note_fixes = [r for r in targets if _normalize_note(r["reviewer_note"]) != r["reviewer_note"]]
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
