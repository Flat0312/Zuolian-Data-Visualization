"""把 phase2 事件证据补证试点的人工审核通过项合并进生产数据。

幂等保证：
- 以 fact_evidences.origin_evidence_id 识别已转正证据；全部已转正时早退不写盘；
- 已转正证据携带的正式 source_id 会复用给同批剩余证据，来源注册前先按 URL 查重；
- 注记类字段追加前均做标记检查，重复执行不产生重复文本或新 ID。

前置条件：
- research/drafts/reports/phase2_event_evidence_pilot_sources.csv 中 10 个有效来源
  （SRC-EVP-005 已确认失效，不合并）已完成网页逐条核验；
- 项目负责人已确认两项决策：EVT-00029 日期降为 1922；EVT-00143 保留"被捕"并加注记。
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
P1_REJECT_MARKER = "由补证试点对应证据"


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


def build_source_reuse_map(
    already_merged: list[dict[str, str]],
    production_rows: list[dict[str, str]],
) -> dict[str, str]:
    """从已转正证据回推试点来源的正式ID，并用URL兜底查重。"""
    reuse: dict[str, str] = {}
    origin_index = {r["origin_evidence_id"]: r for r in production_rows if r["origin_evidence_id"]}
    for row in already_merged:
        merged_row = origin_index.get(row["evidence_id"])
        if merged_row is not None:
            reuse[row["source_id"]] = merged_row["source_id"]
    return reuse


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
    reuse_map = build_source_reuse_map(already_merged, ev_rows)
    url_index = {r["source_url"]: r["source_id"] for r in src_rows if r["source_url"]}
    existing_src_ids = {r["source_id"] for r in src_rows}

    def allocate_source_id() -> str:
        numbers = [int(s.split("-")[1]) for s in existing_src_ids if s.startswith("SRC-")]
        formal_id = f"SRC-{max(numbers) + 1:04d}"
        while formal_id in existing_src_ids:
            numbers.append(max(numbers) + 1)
            formal_id = f"SRC-{max(numbers) + 1:04d}"
        existing_src_ids.add(formal_id)
        return formal_id

    def resolve_source_id(pilot_source_id: str, source_url: str) -> str:
        if pilot_source_id not in reuse_map:
            reuse_map[pilot_source_id] = (
                url_index.get(source_url) if source_url else None
            ) or allocate_source_id()
            if source_url:
                url_index[source_url] = reuse_map[pilot_source_id]
        return reuse_map[pilot_source_id]

    # 1. 注册尚未入库的试点来源（跳过失效来源；URL已存在则复用正式ID）
    registered_sources = 0
    for row in pilot_sources:
        if row["source_id"] in DEPRECATED_SOURCES or row["source_id"] in reuse_map:
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

    # 2. 合并待转正证据
    taken_ids = {r["evidence_id"] for r in ev_rows}
    event_source_updates: dict[str, set[str]] = {}
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
        event_source_updates.setdefault(row["subject_id"], set()).add(row["source_id"])

    # 3. 事件表更新（所有追加均带去重/标记检查）
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

    for row in evt_rows:
        eid = row["event_id"]
        extra_sources = event_source_updates.get(eid)
        if extra_sources:
            current = [s for s in row["source_ids"].split(";") if s]
            row["source_ids"] = ";".join(current + sorted(s for s in extra_sources if s not in current))
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

    # 4. P1 草稿去重标记（带标记检查）
    p1_fields, p1_rows = read_csv(P1_SUPPLEMENT)
    for row in p1_rows:
        if row["evidence_id"] == "FE-SUP-0137" and P1_REJECT_MARKER not in row["reviewer_note"]:
            row["review_status"] = "rejected"
            row["reviewer_note"] = (
                f"2026-08-21 决策：无定位无摘录，{P1_REJECT_MARKER}"
                "（含明确定位与短摘录）替代，本条废弃不合并。" + row["reviewer_note"]
            )

    write_csv(DATA / "sources.csv", src_fields, src_rows)
    write_csv(DATA / "fact_evidences.csv", ev_fields, ev_rows)
    write_csv(DATA / "events.csv", evt_fields, evt_rows)
    write_csv(P1_SUPPLEMENT, p1_fields, p1_rows)

    print(f"sources: +{registered_sources} -> {len(src_rows)}")
    print(f"fact_evidences: +{merged_count} -> {len(ev_rows)}")


if __name__ == "__main__":
    main()
