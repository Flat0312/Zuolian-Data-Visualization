"""把 phase2 第二批补证（龙华二十四烈士名录）合并进生产数据。

幂等保证：
- 以 fact_evidences.origin_evidence_id 识别已转正证据；全部已转正时早退不写盘；
- 已转正证据携带的正式 source_id 会复用给同批剩余证据，来源注册前先按 URL 查重；
- 事件注记与 source_ids 追加均带去重检查，重复执行不产生重复文本或新 ID。

前置条件：两个来源已完成网页核验，项目负责人确认"23姓名+1佚名"展示口径。
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

    _, pilot_sources = read_csv(PILOT_SOURCES)
    _, pilot_evidences = read_csv(PILOT_EVIDENCES)

    origin_index = {r["origin_evidence_id"]: r for r in ev_rows if r["origin_evidence_id"]}
    already_merged = [r for r in pilot_evidences if r["evidence_id"] in origin_index]
    pending = [r for r in pilot_evidences if r["evidence_id"] not in origin_index]

    if not pending:
        print(f"无新增：{len(already_merged)} 条试点证据均已转正，跳过写入。")
        return

    # 来源复用映射：已转正证据回推 + 生产表URL兜底
    reuse_map: dict[str, str] = {}
    for row in already_merged:
        merged_row = origin_index[row["evidence_id"]]
        reuse_map[row["source_id"]] = merged_row["source_id"]
    url_index = {r["source_url"]: r["source_id"] for r in src_rows if r["source_url"]}
    existing_src_ids = {r["source_id"] for r in src_rows}

    def resolve_source_id(pilot_source_id: str, source_url: str) -> str:
        if pilot_source_id not in reuse_map:
            reuse_map[pilot_source_id] = (
                url_index.get(source_url) if source_url else None
            ) or allocate_source_id()
            if source_url:
                url_index[source_url] = reuse_map[pilot_source_id]
        return reuse_map[pilot_source_id]

    def allocate_source_id() -> str:
        numbers = [int(s.split("-")[1]) for s in existing_src_ids if s.startswith("SRC-")]
        formal_id = f"SRC-{max(numbers) + 1:04d}"
        while formal_id in existing_src_ids:
            numbers.append(max(numbers) + 1)
            formal_id = f"SRC-{max(numbers) + 1:04d}"
        existing_src_ids.add(formal_id)
        return formal_id

    # 1. 注册尚未入库的试点来源（URL已存在则复用正式ID）
    # 注意：注册会原地改写 row["source_id"]，须先留存试点ID清单供批次挂接使用。
    pilot_source_ids = [r["source_id"] for r in pilot_sources]
    registered_sources = 0
    for row in pilot_sources:
        if row["source_id"] in reuse_map:
            continue
        if row["source_url"] and row["source_url"] in url_index:
            reuse_map[row["source_id"]] = url_index[row["source_url"]]
            continue
        row["source_id"] = resolve_source_id(row["source_id"], row["source_url"])
        row["needs_manual_review"] = "no"
        if MERGE_NOTE not in row["review_note"]:
            row["review_note"] = f"{row['review_note']}{MERGE_NOTE}。".replace("。。", "。")
        src_rows.append(row)
        registered_sources += 1

    # 本批全部注册来源（含无证据直接引用的交叉核对来源）都必须挂接到目标事件，
    # 否则 schema 会报 orphan_source 警告。
    batch_event_sources = {reuse_map[pid] for pid in pilot_source_ids if pid in reuse_map}

    # 2. 合并待转正证据
    taken_ids = {r["evidence_id"] for r in ev_rows}
    merged_count = 0
    for row in pending:
        row["source_id"] = resolve_source_id(row["source_id"], "")
        original_id = row["evidence_id"]
        row["evidence_id"] = formal_evidence_id(row, taken_ids)
        taken_ids.add(row["evidence_id"])
        row["origin_evidence_id"] = original_id
        row["review_status"] = "reviewed"
        row["reviewer_note"] = f"{MERGE_NOTE}（原编号{original_id}）。{row['reviewer_note']}"
        ev_rows.append(row)
        merged_count += 1

    # 3. 事件表更新（追加均带去重/标记检查）
    for row in evt_rows:
        if row["event_id"] != "EVT-00148":
            continue
        current = [s for s in row["source_ids"].split(";") if s]
        row["source_ids"] = ";".join(current + sorted(s for s in batch_event_sources if s not in current))
        note_append = (
            "2026-08-21 名录口径：官方名录确认23位烈士姓名、另1位佚名烈士，"
            "合葬墓碑刻名22人；名单用字从陵园官网“汤仕佺”。"
        )
        if note_append not in row["display_note"]:
            row["display_note"] = f"{row['display_note']}{note_append}"

    write_csv(DATA / "sources.csv", src_fields, src_rows)
    write_csv(DATA / "fact_evidences.csv", ev_fields, ev_rows)
    write_csv(DATA / "events.csv", evt_fields, evt_rows)

    print(f"sources: +{registered_sources} -> {len(src_rows)}")
    print(f"fact_evidences: +{merged_count} -> {len(ev_rows)}")


if __name__ == "__main__":
    main()
