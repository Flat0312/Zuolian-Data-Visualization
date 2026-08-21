# -*- coding: utf-8 -*-
"""把 phase2 第二批补证（龙华二十四烈士名录）合并进生产数据。

前置条件：两个来源已完成网页核验，项目负责人确认"23姓名+1佚名"展示口径。

动作：
1. sources.csv 追加 SRC-1164（澎湃政务号"龙华英烈"）、SRC-1165（陵园官网英烈名录）。
2. fact_evidences.csv 追加 2 条 FE-EVI-* 证据（review_status=reviewed）。
3. events.csv 更新 EVT-00148：source_ids 追加；display_note 写明名单确认口径。
"""

from __future__ import annotations

import csv
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data" / "processed"
DRAFTS = ROOT / "research" / "drafts" / "reports"

PILOT_SOURCES = DRAFTS / "phase2_batch2_longhua_roster_sources.csv"
PILOT_EVIDENCES = DRAFTS / "phase2_batch2_longhua_roster.csv"
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


def next_source_id(existing: set[str]) -> str:
    numbers = [int(s.split("-")[1]) for s in existing if s.startswith("SRC-")]
    candidate = f"SRC-{max(numbers) + 1:04d}"
    while candidate in existing:
        numbers.append(max(numbers) + 1)
        candidate = f"SRC-{max(numbers) + 1:04d}"
    return candidate


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

    existing_src_ids = {r["source_id"] for r in src_rows}
    id_map: dict[str, str] = {}
    for row in pilot_sources:
        formal_id = next_source_id(existing_src_ids)
        id_map[row["source_id"]] = formal_id
        row["source_id"] = formal_id
        row["needs_manual_review"] = "no"
        row["review_note"] = f"{row['review_note']}{MERGE_NOTE}。".replace("。。", "。")
        src_rows.append(row)
        existing_src_ids.add(formal_id)

    taken_ids = {r["evidence_id"] for r in ev_rows}
    merged_count = 0
    new_sources_for_event: set[str] = set()
    for row in pilot_evidences:
        row["source_id"] = id_map[row["source_id"]]
        original_id = row["evidence_id"]
        row["evidence_id"] = formal_evidence_id(row, taken_ids)
        taken_ids.add(row["evidence_id"])
        row["origin_evidence_id"] = original_id
        row["review_status"] = "reviewed"
        row["reviewer_note"] = f"{MERGE_NOTE}（原编号{original_id}）。{row['reviewer_note']}"
        ev_rows.append(row)
        merged_count += 1
        new_sources_for_event.add(row["source_id"])

    for row in evt_rows:
        if row["event_id"] != "EVT-00148":
            continue
        current = [s for s in row["source_ids"].split(";") if s]
        row["source_ids"] = ";".join(current + sorted(s for s in new_sources_for_event if s not in current))
        note_append = (
            "2026-08-21 名录口径：官方名录确认23位烈士姓名、另1位佚名烈士，"
            "合葬墓碑刻名22人；名单用字从陵园官网“汤仕佺”。"
        )
        if note_append not in row["display_note"]:
            row["display_note"] = f"{row['display_note']}{note_append}"

    write_csv(DATA / "sources.csv", src_fields, src_rows)
    write_csv(DATA / "fact_evidences.csv", ev_fields, ev_rows)
    write_csv(DATA / "events.csv", evt_fields, evt_rows)

    print(f"sources: +{len(id_map)} -> {len(src_rows)}")
    print(f"fact_evidences: +{merged_count} -> {len(ev_rows)}")
    print("id_map:", id_map)


if __name__ == "__main__":
    main()
