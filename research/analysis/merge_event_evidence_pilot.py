# -*- coding: utf-8 -*-
"""把 phase2 事件证据补证试点的人工审核通过项合并进生产数据。

前置条件：
- research/drafts/reports/phase2_event_evidence_pilot_sources.csv 中 10 个有效来源
  （SRC-EVP-005 已确认失效，不合并）已完成网页逐条核验；
- 项目负责人已确认两项决策：EVT-00029 日期降为 1922；EVT-00143 保留"被捕"并加注记。

动作：
1. sources.csv 追加 SRC-1154..SRC-1163（跳过失效来源）。
2. fact_evidences.csv 追加 12 条 FE-EVI-* 正式证据（review_status=reviewed，
   origin_evidence_id 保留原 FE-EVP-* 编号）。
3. events.csv 更新 10 个事件的 source_ids；EVT-00029 日期降级；
   EVT-00143 加术语注记；EVT-00260 置信度 low->medium。
4. P1 草稿中 FE-SUP-0137 标记为被本试点替代，不合并。
"""

from __future__ import annotations

import csv
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"
DRAFTS = ROOT / "research" / "drafts" / "reports"

PILOT_SOURCES = DRAFTS / "phase2_event_evidence_pilot_sources.csv"
PILOT_EVIDENCES = DRAFTS / "phase2_event_evidence_pilot.csv"
P1_SUPPLEMENT = DRAFTS / "phase1_p1_evidence_supplement.csv"

DEPRECATED_SOURCES = {"SRC-EVP-005"}
MERGE_NOTE = "2026-08-21 合并转正：网页逐条核验通过（AI辅助），项目负责人确认合并"


def read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        return list(reader.fieldnames or []), list(reader)


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def next_source_id(existing: list[str]) -> str:
    numbers = [int(s.split("-")[1]) for s in existing if s.startswith("SRC-")]
    return f"SRC-{max(numbers) + 1:04d}"


def formal_evidence_id(row: dict[str, str], taken: set[str]) -> str:
    basis = "|".join(
        row[key] for key in ("subject_type", "subject_id", "predicate", "object_value", "source_id")
    )
    digest = hashlib.md5(basis.encode("utf-8")).hexdigest()[:10].upper()
    candidate = f"FE-EVI-{digest}"
    while candidate in taken:
        digest = hashlib.md5((candidate + "x").encode("utf-8")).hexdigest()[:10].upper()
        candidate = f"FE-EVI-{digest}"
    return candidate


def main() -> None:
    src_fields, src_rows = read_csv(DATA / "sources.csv")
    ev_fields, ev_rows = read_csv(DATA / "fact_evidences.csv")
    evt_fields, evt_rows = read_csv(DATA / "events.csv")

    pilot_src_fields, pilot_sources = read_csv(PILOT_SOURCES)
    pilot_ev_fields, pilot_evidences = read_csv(PILOT_EVIDENCES)

    # 1. 来源合并：跳过失效来源，分配正式 ID
    existing_src_ids = {r["source_id"] for r in src_rows}
    id_map: dict[str, str] = {}
    for row in pilot_sources:
        if row["source_id"] in DEPRECATED_SOURCES:
            continue
        formal_id = next_source_id([r["source_id"] for r in src_rows])
        while formal_id in existing_src_ids:
            formal_id = next_source_id([r["source_id"] for r in src_rows] + list(id_map.values()))
        id_map[row["source_id"]] = formal_id
        row["source_id"] = formal_id
        row["needs_manual_review"] = "no"
        row["review_note"] = f"{row['review_note']}{MERGE_NOTE}。".replace("。。", "。")
        src_rows.append(row)
        existing_src_ids.add(formal_id)

    # 2. 证据合并：生成正式 ID，标记 reviewed，保留原编号
    taken_ids = {r["evidence_id"] for r in ev_rows}
    event_source_updates: dict[str, set[str]] = {}
    merged_count = 0
    for row in pilot_evidences:
        old_source = row["source_id"]
        row["source_id"] = id_map[old_source]
        original_id = row["evidence_id"]
        row["evidence_id"] = formal_evidence_id(row, taken_ids)
        taken_ids.add(row["evidence_id"])
        row["origin_evidence_id"] = original_id
        row["review_status"] = "reviewed"
        row["reviewer_note"] = f"{MERGE_NOTE}（原编号{original_id}）。{row['reviewer_note']}"
        ev_rows.append(row)
        merged_count += 1
        event_source_updates.setdefault(row["subject_id"], set()).add(row["source_id"])

    # 3. 事件表更新
    decisions = {
        "EVT-00029": {
            "event_date": "1922",
            "date_precision": "年",
            "canonical_event_key": None,
            "confidence": None,
            "needs_manual_review": "no",
            "display_note": (
                "1922年，丁玲进入平民女校求学（该校创办于1922年2月），"
                "当前条目按年份精度展示；抵沪具体月份证据不足，已由1922-02降级。"
            ),
            "correction_reason_append": "2026-08-21 补证试点：证据仅支持在校年份，日期精度由月降为年。",
        },
        "EVT-00143": {
            "display_note_append": "权威来源原始表述为“秘密绑架”，正式名称保留“被捕”以统一术语。",
        },
        "EVT-00260": {
            "confidence": "medium",
            "correction_reason_append": "2026-08-21 补证试点：获得B级可定位来源支持，置信度由low升为medium。",
        },
    }

    updated_events = []
    for row in evt_rows:
        eid = row["event_id"]
        extra_sources = event_source_updates.get(eid)
        if extra_sources:
            current = [s for s in row["source_ids"].split(";") if s]
            merged = current + sorted(s for s in extra_sources if s not in current)
            row["source_ids"] = ";".join(merged)
        if eid in decisions:
            rule = decisions[eid]
            for key in ("event_date", "date_precision", "confidence", "needs_manual_review"):
                if rule.get(key):
                    row[key] = rule[key]
            if rule.get("canonical_event_key") is None and rule.get("event_date"):
                parts = row["canonical_event_key"].split("|")
                parts[-1] = rule["event_date"]
                row["canonical_event_key"] = "|".join(parts)
            if rule.get("display_note"):
                row["display_note"] = rule["display_note"]
            if rule.get("display_note_append") and rule["display_note_append"] not in row["display_note"]:
                row["display_note"] = f"{row['display_note']}{rule['display_note_append']}"
            if rule.get("correction_reason_append") and rule["correction_reason_append"] not in row["correction_reason"]:
                row["correction_reason"] = f"{row['correction_reason']}{rule['correction_reason_append']}"
            updated_events.append(eid)
        elif extra_sources:
            updated_events.append(eid)

    # 4. P1 草稿去重标记
    p1_fields, p1_rows = read_csv(P1_SUPPLEMENT)
    for row in p1_rows:
        if row["evidence_id"] == "FE-SUP-0137":
            row["review_status"] = "rejected"
            row["reviewer_note"] = (
                "2026-08-21 决策：无定位无摘录，由补证试点对应证据"
                "（含明确定位与短摘录）替代，本条废弃不合并。" + row["reviewer_note"]
            )

    write_csv(DATA / "sources.csv", src_fields, src_rows)
    write_csv(DATA / "fact_evidences.csv", ev_fields, ev_rows)
    write_csv(DATA / "events.csv", evt_fields, evt_rows)
    write_csv(P1_SUPPLEMENT, p1_fields, p1_rows)

    print(f"sources: +{len(id_map)} -> {len(src_rows)}")
    print(f"fact_evidences: +{merged_count} -> {len(ev_rows)}")
    print(f"events updated: {sorted(set(updated_events))}")
    print("id_map:", id_map)


if __name__ == "__main__":
    main()
